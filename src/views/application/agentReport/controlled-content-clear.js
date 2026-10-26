import { ICommandService, IUndoRedoService } from '@univerjs/core'
import {
  ClearSelectionContentCommand,
  RangeProtectionPermissionEditPoint,
  SetRangeValuesMutation,
  SetRangeValuesUndoMutationFactory,
  SheetInterceptorService,
  SheetPermissionCheckController,
  WorkbookEditablePermission,
  WorksheetEditPermission,
  WorksheetSetCellValuePermission
} from '@univerjs/sheets'

/** Prepare the SDK content mutation used by the existing clearRange handler.
 * The high-level clear command also changes auto-height. Do not execute it and
 * do not repair/ignore its extra changes afterwards. All plans are prepared
 * before executeActions starts, so an unavailable/unsafe clear is zero-write.
 */
export function prepareControlledContentClear(sheet, range) {
  if (typeof sheet.getInject !== 'function' || typeof sheet.getWorkbook !== 'function') throw new Error('安全的内容清除能力不可用，拒绝清除')
  const injector = sheet.getInject()
  const unitId = sheet.getWorkbook()?.getUnitId?.()
  const subUnitId = sheet.getSheetId?.()
  if (!unitId || !subUnitId || !injector || typeof injector.invoke !== 'function') throw new Error('内容清除目标无法核实')
  return injector.invoke(accessor => {
    const commands = accessor.get(ICommandService)
    const undo = accessor.get(IUndoRedoService)
    const permissions = accessor.get(SheetPermissionCheckController)
    const interceptors = accessor.get(SheetInterceptorService)
    if (!commands.hasCommand(SetRangeValuesMutation.id) || typeof undo.pushUndoRedo !== 'function') throw new Error('内容清除命令或撤销能力不可用')
    const permissionTypes = {
      workbookTypes: [WorkbookEditablePermission],
      worksheetTypes: [WorksheetSetCellValuePermission, WorksheetEditPermission],
      rangeTypes: [RangeProtectionPermissionEditPoint]
    }
    const checkPermission = () => {
      if (permissions.permissionCheckWithRanges(permissionTypes, [range], unitId, subUnitId) !== true) throw new Error('没有目标范围的内容修改权限')
    }
    checkPermission()
    // Do not silently skip plugin effects (e.g. dependent resources). This
    // limited content path opens only when there are no additional effects.
    const effects = interceptors.onCommandExecute({ id: ClearSelectionContentCommand.id, params: { unitId, subUnitId, ranges: [range] } })
    if (!effects || !Array.isArray(effects.redos) || !Array.isArray(effects.undos) || effects.redos.length || effects.undos.length) throw new Error('清除涉及尚未支持的关联变更，拒绝执行')
    const cellValue = {}
    for (let row = range.startRow; row <= range.endRow; row++) {
      cellValue[row] = {}
      for (let column = range.startColumn; column <= range.endColumn; column++) cellValue[row][column] = { v: null, p: null, f: null, si: null }
    }
    const params = { unitId, subUnitId, cellValue }
    const previous = SetRangeValuesUndoMutationFactory(accessor, params)
    return () => {
      checkPermission()
      // Use the SDK mutation, including its normal command notifications and
      // cell merge semantics; never edit the workbook model or rowData directly.
      if (commands.syncExecuteCommand(SetRangeValuesMutation.id, params) !== true) throw new Error('内容清除未确认执行成功')
      undo.pushUndoRedo({
        unitID: unitId,
        undoMutations: [{ id: SetRangeValuesMutation.id, params: previous }],
        redoMutations: [{ id: SetRangeValuesMutation.id, params }]
      })
    }
  })
}

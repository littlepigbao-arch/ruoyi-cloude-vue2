'use strict';
// In-memory SDK boundaries for unit tests; never a real Univer/DB acceptance claim.
const core={ICommandService:'commands',IUndoRedoService:'undo'};
const names=['SheetPermissionCheckController','SheetInterceptorService','WorkbookEditablePermission','WorksheetSetCellValuePermission','WorksheetEditPermission','RangeProtectionPermissionEditPoint'];
const sheets=Object.fromEntries(names.map(n=>[n,n]));
sheets.ClearSelectionContentCommand={id:'sheet.command.clear-selection-content'};
sheets.SetRangeValuesMutation={id:'sheet.mutation.set-range-values'};
sheets.SetRangeValuesUndoMutationFactory=(_accessor,params)=>JSON.parse(JSON.stringify(params));
function attach(sheet, options={}){
 const history=[],sent=[];
 const services={
  SheetPermissionCheckController:{permissionCheckWithRanges:()=>options.permitted!==false},
  SheetInterceptorService:{onCommandExecute:()=>options.intercepted||{redos:[],undos:[]}},
  undo:{pushUndoRedo:entry=>{if(options.undoThrows)throw Error('undo failed');history.push(entry);}},
  commands:{hasCommand:()=>!options.missingMutation,syncExecuteCommand:(id,params)=>{
   sent.push({id,params});if(options.mutationThrows)throw Error('mutation failed');if(Object.hasOwn(options,'mutationResult')&&options.mutationResult!==true)return options.mutationResult;
   if(options.mutate)options.mutate(params);else for(const [row,cols]of Object.entries(params.cellValue))for(const column of Object.keys(cols))sheet.getRange(Number(row),Number(column),1,1).clearContent();
   return true;
  }}
 };
 const accessor={get:key=>services[key]},injector={invoke:fn=>fn(accessor)};
 sheet.getInject=()=>injector;sheet.getWorkbook=()=>({getUnitId:()=> 'unit'});
 return {history,sent,services};
}
module.exports={modules:{'@univerjs/core':core,'@univerjs/sheets':sheets},attach};

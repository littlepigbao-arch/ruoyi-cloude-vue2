/** A1: construct every draft save from the version actually loaded by the page. */
export function workingStatePayload(state, snapshot) {
  if (!state.loadedRevision || !Number.isInteger(state.draftVersion)) throw new Error('页面尚未核实草稿版本，请保留编辑并重新加载')
  return {
    branchId: state.currentBranchId, baseCommitId: state.headCommitId, baseVersion: state.branchVersion,
    expectedDraftVersion: state.draftVersion, loadedRevision: state.loadedRevision,
    content: JSON.stringify(snapshot), plannedActions: state.pendingActions, executedChanges: state.pendingChanges
  }
}

export function sameWorkspace(state, identity) {
  return String(state.currentDocumentId) === String(identity.workbookId) &&
    String(state.currentBranchId) === String(identity.branchId) && state.workspaceGeneration === identity.generation
}

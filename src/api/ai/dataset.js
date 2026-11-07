import { agentRequest } from './agent'

// C1a chunked artifact reads. Cursor is keyset-based: pass back the returned nextCursor; null means the last page.
// Never fetch all pages at once; consume one chunk per request to keep payloads bounded.
export const getArtifact = id => agentRequest('/artifacts/' + encodeURIComponent(id), null, 'get')
export const getArtifactChunks = (id, cursor) =>
  agentRequest('/artifacts/' + encodeURIComponent(id) + '/chunks' + (cursor === undefined || cursor === null ? '' : '?cursor=' + encodeURIComponent(cursor)), null, 'get')

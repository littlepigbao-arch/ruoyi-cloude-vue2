import request from '@/utils/request'

export async function agentRequest(path, data, method = 'post') {
  const response = await request({ url: '/ai/agent' + path, method, data, timeout: 60000 })
  if (!response || response.code !== 200) throw new Error((response && response.msg) || '智能体请求失败')
  return response.data
}
export const agentCapabilities = () => agentRequest('/capabilities', null, 'get')
export const agentWorkspace = id => agentRequest('/workspace/' + id, null, 'get')
export const activeAgent = id => agentRequest('/runs/active?workbookId=' + encodeURIComponent(id), null, 'get')
export const getAgent = id => agentRequest('/runs/' + id, null, 'get')

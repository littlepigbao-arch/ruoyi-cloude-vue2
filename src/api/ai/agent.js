import axios from 'axios'
import { getToken } from '@/utils/auth'

// Stable server request/call IDs provide idempotency. Do not suppress an identical receipt retry.
const agentHttp = axios.create({ baseURL: process.env.VUE_APP_BASE_API, timeout: 60000 })

export async function agentRequest(path, data, method = 'post') {
  const token = getToken()
  const result = await agentHttp.request({ url: '/ai/agent' + path, method, data,
    headers: { 'Content-Type': 'application/json', Authorization: token ? 'Bearer ' + token : '' } })
  const response = result.data
  if (!response || response.code !== 200) {
    const error = new Error((response && response.msg) || '智能体请求失败')
    error.code = response && response.code
    error.definitive = error.code === 409 || error.code === 422
    throw error
  }
  return response.data
}
export const agentCapabilities = () => agentRequest('/capabilities', null, 'get')
export const agentWorkspace = id => agentRequest('/workspace/' + id, null, 'get')
export const activeAgent = id => agentRequest('/runs/active?workbookId=' + encodeURIComponent(id), null, 'get')
export const getAgent = id => agentRequest('/runs/' + id, null, 'get')

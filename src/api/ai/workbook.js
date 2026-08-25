/**
 * AI 智能报表文档接口封装
 * 使用独立 axios 实例：文档快照 JSON 体积可能较大，需更长超时，
 * 且避开 src/utils/request.js 的 10s 超时、5M 防重复提交限制与全局错误弹窗。
 */
import axios from 'axios'
import { getToken } from '@/utils/auth'

const wbHttp = axios.create({
  baseURL: process.env.VUE_APP_BASE_API, // /dev-api
  timeout: 120000, // 120s，容纳大 JSON 上传
  headers: { 'Content-Type': 'application/json;charset=utf-8' }
})

// 请求拦截器：注入 token
wbHttp.interceptors.request.use(
  (config) => {
    const token = getToken()
    if (token) config.headers['Authorization'] = 'Bearer ' + token
    return config
  },
  (error) => Promise.reject(error)
)

// 响应拦截器：直接返回 res.data（{code,msg,data}）
wbHttp.interceptors.response.use(
  (res) => res.data,
  (error) => Promise.reject(error)
)

/**
 * 查询当前用户的文档列表（不含内容）
 * @returns {Promise<{code,msg,data:Array}>}
 */
export function listWorkbooks() {
  return wbHttp.get('/ai/workbook/list')
}

/**
 * 查询单条文档（含完整快照内容）
 * @param {number|string} workbookId
 * @returns {Promise<{code,msg,data}>}
 */
export function getWorkbook(workbookId) {
  return wbHttp.get('/ai/workbook/' + workbookId)
}

/**
 * 保存/更新文档
 * @param {object} payload {workbookId?, name, type, content}
 * @returns {Promise<{code,msg,data}>}
 */
export function saveWorkbook(payload) {
  return wbHttp.post('/ai/workbook/save', payload)
}

/**
 * 删除文档
 * @param {number|string} workbookId
 * @returns {Promise<{code,msg,data}>}
 */
export function deleteWorkbook(workbookId) {
  return wbHttp.delete('/ai/workbook/' + workbookId)
}

/**
 * 重命名文档
 * @param {object} payload {workbookId, name}
 * @returns {Promise<{code,msg,data}>}
 */
export function renameWorkbook(payload) {
  return wbHttp.put('/ai/workbook/rename', payload)
}

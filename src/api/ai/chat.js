/**
 * AI 对话接口封装
 * 关键：使用独立 axios 实例，避开 src/utils/request.js 的 10s 超时和全局 401/500/601 弹窗
 * 后端未就绪或网络异常时自动回退 mock 解析器，保证可演示
 */
import axios from 'axios'
import { getToken } from '@/utils/auth'
import { mockParse } from '@/views/application/agentReport/ai-mock-parser'

// 独立 axios 实例（不复用 request.js）
const aiHttp = axios.create({
  baseURL: process.env.VUE_APP_BASE_API, // /dev-api
  timeout: 60000, // 60s，覆盖 request.js 的 10s
  headers: { 'Content-Type': 'application/json;charset=utf-8' }
})

// 请求拦截器：注入 token
aiHttp.interceptors.request.use(
  (config) => {
    const token = getToken()
    if (token) config.headers['Authorization'] = 'Bearer ' + token
    return config
  },
  (error) => Promise.reject(error)
)

// 响应拦截器：直接返回 res.data，不挂全局弹窗
aiHttp.interceptors.response.use(
  (res) => res.data,
  (error) => Promise.reject(error)
)

/**
 * 发起 AI 对话
 * @param {object} payload {message, conversationId, history, context, clientMeta}
 * @param {object} opts {useMock:bool, allowMockFallback:bool}
 * @returns {Promise<{code,msg,data,source?}>}
 */
export async function sendChat(payload, opts = {}) {
  // 1. 本地模式：强制走 mock
  if (opts.useMock) {
    return mockToResponse(payload, mockParse(payload.message, payload.context), 'mock')
  }
  // 2. 远程模式
  try {
    const data = await aiHttp.post('/ai/chat', payload)
    if (data && data.code === 200) {
      // AI 拒绝兜底：AI 返回空 actions 且回复含"不支持筛选/FILTER 函数被禁用"等关键词
      // 时，说明 AI 不知道前端已支持该指令。本地二次解析覆盖。
      const aiActions = (data.data && data.data.actions) || []
      const aiReply = (data.data && data.data.reply) || ''
      if (aiActions.length === 0 && looksLikeRejectedByAi(aiReply) && payload.message) {
        const fallback = mockParse(payload.message, payload.context)
        if (fallback.actions && fallback.actions.length > 0) {
          return mockToResponse(payload, fallback, 'ai-rejected-local-fallback')
        }
      }
      return data
    }
    // 503 服务不可用 → 回退 mock
    if (data && data.code === 503 && opts.allowMockFallback) {
      return mockToResponse(payload, mockParse(payload.message, payload.context), 'mock-fallback')
    }
    // 其他业务错误 → 抛错让调用方处理
    const err = new Error((data && data.msg) || 'AI 接口返回异常')
    err.code = data && data.code
    throw err
  } catch (e) {
    // 网络异常/超时/404（后端未就绪）→ 回退 mock
    if (opts.allowMockFallback && isNetworkError(e)) {
      return mockToResponse(payload, mockParse(payload.message, payload.context), 'mock-fallback')
    }
    throw e
  }
}

/**
 * 查询某会话的历史消息
 * @param {string} conversationId
 * @returns {Promise<{code,msg,data:Array}>} data 为 [{role,content,...}] 按时间升序
 */
export async function getHistory(conversationId) {
  const data = await aiHttp.get('/ai/history/' + encodeURIComponent(conversationId))
  return data
}

/**
 * 检测 AI 回复是否表明它"拒绝了"该指令（误判前端能力）
 * 触发条件：回复文本含特定关键词，且用户请求本质是表格操作
 */
function looksLikeRejectedByAi(reply) {
  if (!reply || typeof reply !== 'string') return false
  const r = reply
  // 关键词：AI 自承"不支持/暂不支持" + 筛选相关字眼
  const hasReject = /暂不支持|不支持|无法实现|前端禁用|无法直接|不能直接/.test(r)
  const hasFilterHint = /筛选|FILTER|过滤|COUNTIF|Ctrl\+Shift/.test(r)
  return hasReject && hasFilterHint
}

/**
 * 判断是否为可回退的网络错误
 */
function isNetworkError(e) {
  if (!e) return false
  const msg = String(e.message || '')
  return (
    msg.indexOf('Network Error') >= 0 ||
    msg.indexOf('timeout') >= 0 ||
    msg.indexOf('404') >= 0 ||
    msg.indexOf('Request failed with status code 5') >= 0 ||
    msg.indexOf('ERR_CONNECTION') >= 0
  )
}

/**
 * 把 mock 解析结果包装成与后端一致的响应结构
 */
function mockToResponse(payload, parsed, source) {
  return {
    code: 200,
    msg: 'ok',
    source: source,
    data: {
      reply: parsed.reply,
      conversationId: payload.conversationId || randomUuid(),
      actions: parsed.actions,
      needFeedback: false
    }
  }
}

function randomUuid() {
  return 'mock-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8)
}

export default {
  sendChat
}

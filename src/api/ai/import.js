import axios from 'axios'
import { getToken } from '@/utils/auth'

// C3-FE-02 import job client (backend: WorkbookImportController, /ai/agent/import/*).
// createImportJob is the only multipart call: requestId is the server-side idempotency key and the
// file travels as FormData so the runtime sets the multipart boundary itself - a hand-written
// Content-Type (especially the JSON one agent.js uses) would corrupt the boundary. The GET entries
// mirror agent.js error semantics exactly: code!==200 throws with error.code / error.definitive
// (409 carries the structured IMPORT_* refusal whose payload embeds kind/limit/observed).
// UI wiring (upload split / progress) lands in C4; this module is API-only.
//
// Inherited-defaults hazard (F1): src/utils/request.js:14 sets the global
// axios.defaults.headers['Content-Type'] = 'application/json;charset=utf-8' before this module is
// evaluated and axios.create() snapshots it. With that header present, the installed axios converts
// FormData into a JSON string (defaults/index.js), replacing the file bytes. importTransformRequest
// is the deletion point: it drops the header for FormData payloads only (the multipart boundary
// must stay runtime-generated) and delegates every other payload to axios.defaults.transformRequest
// so plain JSON bodies keep their previous behavior.
//
// Precision contract (F2 + Codex 2nd round): the default transformResponse materializes every JSON
// number token as a JS Number. importTransformResponse keeps the exact original literal as a string
// for tokens whose decimal value a Number cannot carry losslessly - not only >17-significant-digit
// decimals but also out-of-range exponents (1e400 -> Infinity, 1E-400 -> 0) and subnormal rounding
// (4e-324 -> 5e-324; even 15 significant digits lose relative precision near the subnormal edge,
// e.g. 1.23456789012345e-310) - while every value-equal Number (123 / -45.75 / 45838.99998842592 /
// 1e2 == 100) stays a JS Number (C4 routes the string literals through the Univer text lane).
// The lossy-literal parking spot is a single-key object keyed by a per-parse nonce, never a string
// prefix: a genuine payload string wearing a sentinel-looking prefix must survive verbatim.

// Mirror of the installed axios utils.isFormData (kept local to avoid depending on internals):
// Node/Browser FormData, cross-realm instances and toString shims all qualify.
function isFormData(data) {
  return Boolean(data) && (
    (typeof FormData === 'function' && data instanceof FormData) ||
    Object.prototype.toString.call(data) === '[object FormData]' ||
    (typeof data.toString === 'function' && data.toString() === '[object FormData]')
  )
}

function importTransformRequest(data, headers) {
  if (isFormData(data)) {
    // dispatchRequest runs this transform before header flattening and hands this exact headers
    // object to the adapter, so the delete propagates; without it the inherited JSON header makes
    // axios stringify the FormData and destroy the file bytes.
    delete headers['Content-Type']
    return data
  }
  return axios.defaults.transformRequest.reduce((d, fn) => fn.call(this, d, headers), data)
}

// JSON number grammar (RFC 8259); sticky so the scanner only ever matches at a token boundary.
const NUMBER_TOKEN = /-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/y
// Canonical decimal text: optional sign, digits, optional fraction, optional exponent. Matches
// every JSON number token and every String(finite Number) form ('1e+21', '1e-7', '-45.75', '0').
const DECIMAL_TEXT = /^(-?)(\d+)(?:\.(\d+))?(?:[eE]([+-]?\d+))?$/

// Exact decimal value of canonical numeric text as { mant, exp } with value = mant * 10^exp:
// mant is the signed digit string without the decimal point (leading zeros stripped), exp is the
// literal exponent minus the fraction digit count. BigInt keeps every comparison exact. Returns
// null for non-canonical text - defensive: the caller then keeps the literal as a string, never
// guessing a Number.
function decimalParts(text) {
  const m = DECIMAL_TEXT.exec(text)
  if (!m) return null
  const frac = m[3] || ''
  return {
    mant: BigInt(m[1] + (m[2] + frac).replace(/^0+(?=\d)/, '')),
    exp: (m[4] ? Number(m[4]) : 0) - frac.length
  }
}

// 10 ** k by loop multiplication: with a nonzero normalized mantissa and a finite Number, |exp| is
// bounded by ~325 + significant-digit count (values outside [4.94e-324, 1.8e308] are killed by the
// Infinity/zero short-circuits before this runs), so the loop stays tiny for any real token.
function pow10(k) {
  let out = 1n
  for (let i = 0; i < k; i++) out *= 10n
  return out
}

// Value equality of two exact decimals at the common exponent E = min(exp1, exp2): the mantissas
// scaled into that exponent must agree (exact BigInt cross-multiplication, no float involved).
// A zero mantissa equals only another zero, which also folds -0 into the Number lane.
function decimalValueEquals(a, b) {
  if (!a || !b) return false
  if (a.mant === 0n || b.mant === 0n) return a.mant === b.mant
  const common = Math.min(a.exp, b.exp)
  return a.mant * pow10(a.exp - common) === b.mant * pow10(b.exp - common)
}

// A token may become a Number only when its exact decimal value equals the exact decimal value of
// String(Number(token)) - value equivalence, not format equivalence (1.0 -> 1 and 1e2 -> 100 stay
// numeric; the spelled digits themselves are irrelevant).
// The old <=15-significant-digit fast path was unsafe and is gone: relative precision degrades in
// the subnormal region (15 digits still round near 1e-310), and out-of-range exponents overflow to
// Infinity or underflow to 0. A byte-wise String(n) comparison alone misses those cases from the
// other side, so both texts are parsed into exact decimals and compared value by value.
function isLosslessInNumber(token) {
  const n = Number(token)
  if (!Number.isFinite(n)) return false // 1e400: the Number side is Infinity, never value-equal
  return decimalValueEquals(decimalParts(token), decimalParts(String(n)))
}

// Outside JSON strings, replace number tokens whose value a Number would round with a single-key
// JSON object {"<rawKey>":"<token>"} so the subsequent JSON.parse parks the literal intact;
// strings and their digit-like text are left verbatim. Both the key and the value MUST go through
// JSON.stringify: the key embeds NUL characters, which are illegal unescaped inside JSON string
// text (only the \u0000 escape form is valid).
function guardLossyNumbers(text, rawKey) {
  let out = '', inString = false
  for (let i = 0; i < text.length;) {
    const ch = text[i]
    if (inString) {
      if (ch === '\\') { out += text.slice(i, i + 2); i += 2; continue }
      inString = ch !== '"'
      out += ch
      i++
      continue
    }
    if (ch === '"') { inString = true; out += ch; i++; continue }
    if (ch === '-' || (ch >= '0' && ch <= '9')) {
      NUMBER_TOKEN.lastIndex = i
      const token = NUMBER_TOKEN.exec(text)
      if (token) {
        out += isLosslessInNumber(token[0])
          ? token[0]
          : '{' + JSON.stringify(rawKey) + ':' + JSON.stringify(token[0]) + '}'
        i = NUMBER_TOKEN.lastIndex
        continue
      }
    }
    out += ch
    i++
  }
  return out
}

// Unwrap only the wrapper objects this parse created: a plain object whose ONLY key is this
// parse's nonce key restores to the parked literal string. Collision analysis (why the old
// '\u0000RAW\u0000' string prefix was retired): strings are returned verbatim, so a genuine
// payload string wearing any sentinel-looking prefix (e.g. the POI shared-string decode product
// of _x0000_RAW_x0000_...) can never be touched; and an object could only be mistaken for a
// wrapper if a real payload happened to own exactly the one nonce key of this parse - the nonce
// is fresh random per parse, so that is negligible. Arrays and other objects recurse.
function restoreGuardedNumbers(value, rawKey) {
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    const keys = Object.keys(value)
    if (keys.length === 1 && keys[0] === rawKey) return value[rawKey]
    for (const key of keys) value[key] = restoreGuardedNumbers(value[key], rawKey)
    return value
  }
  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) value[i] = restoreGuardedNumbers(value[i], rawKey)
  }
  return value
}

// Each parse gets its own nonce key ('\u0000RAW\u0000' + random base36), so literals parked by
// this call can only be unwrapped by this same call: parallel or nested parses never
// cross-contaminate, and any stale wrapper would simply stay an ordinary object.
function parsePreservingNumbers(text) {
  const rawKey = '\u0000RAW\u0000' + Math.random().toString(36).slice(2, 12)
  return restoreGuardedNumbers(JSON.parse(guardLossyNumbers(text, rawKey)), rawKey)
}

function importTransformResponse(data) {
  if (typeof data !== 'string') return data
  try {
    return parsePreservingNumbers(data)
  } catch (e) {
    // Mirror axios silentJSONParsing: a malformed body falls through untouched to the caller.
    return data
  }
}

const importHttp = axios.create({
  baseURL: process.env.VUE_APP_BASE_API,
  timeout: 60000,
  transformRequest: [importTransformRequest],
  transformResponse: [importTransformResponse]
})

async function importRequest(config) {
  const token = getToken()
  const result = await importHttp.request({
    ...config,
    headers: { ...(config.headers || {}), Authorization: token ? 'Bearer ' + token : '' }
  })
  const response = result.data
  if (!response || response.code !== 200) {
    const error = new Error((response && response.msg) || '导入请求失败')
    error.code = response && response.code
    error.definitive = error.code === 409 || error.code === 422
    throw error
  }
  return response.data
}

// requestId is the idempotency key: re-posting the same key never creates a second job.
export const createImportJob = (file, requestId) => {
  const form = new FormData()
  if (file && file.name) form.append('file', file, file.name)
  else form.append('file', file)
  form.append('requestId', requestId)
  return importRequest({ url: '/ai/agent/import/jobs', method: 'post', data: form })
}

export const getImportJob = id =>
  importRequest({ url: '/ai/agent/import/jobs/' + encodeURIComponent(id), method: 'get' })

export const getImportJobByRequest = requestId =>
  importRequest({ url: '/ai/agent/import/jobs?requestId=' + encodeURIComponent(requestId), method: 'get' })

// sourceHash must be the hash bound at create time (server rejects a preview against a foreign
// hash); rows defaults to the backend's bounded default of 50.
export const previewImport = (id, sourceHash, rows = 50) =>
  importRequest({
    url: '/ai/agent/import/jobs/' + encodeURIComponent(id) + '/preview?sourceHash=' + encodeURIComponent(sourceHash) + '&rows=' + rows,
    method: 'get'
  })

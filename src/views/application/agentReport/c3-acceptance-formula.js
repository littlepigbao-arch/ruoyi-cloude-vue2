'use strict'
// C3 acceptance harness formula evaluator (pure function module, no Vue/SDK/HTTP dependencies).
//
// Why this exists: the harness page's formula verdict used to be "the computed result is SOME
// finite number", which let any number - including the fixture's deliberately stale cached 13 -
// pass as a false green. The real verdict must be an exact value match: the page replays the
// preview rows into Univer, reads the recalculated cell value, and compares it (double ===
// double, same JS arithmetic on both sides) against the expectation this module computes from the
// formula text and the replayed reference values.
//
// Scope is deliberately minimal: ONLY simple four-arithmetic formulas over cell references and
// numeric literals - the whitelist below. Anything else (function names, strings, ranges,
// percent, comparison operators, ...) is NOT decidable here and yields null; the caller must
// record such cells as 人工核对 (manual check) and must never guess a verdict. That is why there
// is no eval() and no partial parsing: one whitelist scan up front rejects the whole formula.

// Whitelist scan: the whole text must be composed of nothing but cell references (A1 / $B$2 /
// $a1 ... optional $ before column and row, 1-3 letters, 1-7 digits), numeric literals (digits
// with optional fraction and exponent) and the operators + - * / ( ). The scan runs on the
// whitespace-stripped text purely as a gate; the tokenizer below still reads the ORIGINAL text so
// 'A 1' and '=A1 1' (a space gluing two tokens) stay undecidable. A failing scan means
// "undecidable" - the caller records 人工核对, never a numeric verdict.
const SIMPLE_FORMULA = /^(?:\$?[A-Za-z]{1,3}\$?[0-9]{1,7}|[0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?|[+\-*/()])+$/

// One token at a time; sticky regexes so the scanner only ever advances at token boundaries.
const TOKEN_REF = /\$?[A-Za-z]{1,3}\$?[0-9]{1,7}/y
const TOKEN_NUMBER = /[0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?/y

// "A1" -> 0-based column index (A=0, B=1, ... Z=25, AA=26 ...).
function columnIndexOf(letters) {
  let column = 0
  for (let i = 0; i < letters.length; i++) column = column * 26 + (letters.charCodeAt(i) - 64)
  return column - 1
}

// Recursive-descent over the token stream with the usual precedence:
//   expr   := term (('+' | '-') term)*
//   term   := factor (('*' | '/') factor)*
//   factor := ('+' | '-') factor | primary
//   primary:= NUMBER | CELLREF | '(' expr ')'
// resolveCell receives the normalized reference (absolute $ stripped, uppercased, e.g. 'A1') and
// must return a finite number; anything else (null/undefined/text) makes the whole formula null.
function evaluate(text, resolveCell) {
  const tokens = []
  for (let i = 0; i < text.length;) {
    if (/[\s]/.test(text[i])) { i++; continue }
    if (text[i] === '$' || /[A-Za-z]/.test(text[i])) {
      TOKEN_REF.lastIndex = i
      const match = TOKEN_REF.exec(text)
      if (!match) return null
      const raw = match[0]
      const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(raw)
      if (!m) return null
      tokens.push({ kind: 'ref', ref: m[1].toUpperCase() + m[2] })
      i = TOKEN_REF.lastIndex
      continue
    }
    if (/[0-9]/.test(text[i])) {
      TOKEN_NUMBER.lastIndex = i
      const match = TOKEN_NUMBER.exec(text)
      if (!match) return null
      tokens.push({ kind: 'num', value: Number(match[0]) })
      i = TOKEN_NUMBER.lastIndex
      continue
    }
    if ('+-*/()'.includes(text[i])) { tokens.push({ kind: 'op', value: text[i] }); i++; continue }
    return null // any other character (comma, quote, percent, <, &, ...) -> undecidable
  }
  if (!tokens.length) return null

  let pos = 0
  const peek = () => tokens[pos]
  const expr = () => {
    let left = term()
    if (left === null) return null
    while (peek() && peek().kind === 'op' && (peek().value === '+' || peek().value === '-')) {
      const op = tokens[pos++].value
      const right = term()
      if (right === null) return null
      left = op === '+' ? left + right : left - right
    }
    return left
  }
  const term = () => {
    let left = factor()
    if (left === null) return null
    while (peek() && peek().kind === 'op' && (peek().value === '*' || peek().value === '/')) {
      const op = tokens[pos++].value
      const right = factor()
      if (right === null) return null
      left = op === '*' ? left * right : left / right
    }
    return left
  }
  const factor = () => {
    const token = peek()
    if (token && token.kind === 'op' && (token.value === '+' || token.value === '-')) {
      pos++
      const inner = factor()
      if (inner === null) return null
      return token.value === '-' ? -inner : inner
    }
    return primary()
  }
  const primary = () => {
    const token = peek()
    if (!token) return null
    if (token.kind === 'num') { pos++; return token.value }
    if (token.kind === 'ref') {
      pos++
      const value = resolveCell(token.ref)
      return typeof value === 'number' && Number.isFinite(value) ? value : null
    }
    if (token.kind === 'op' && token.value === '(') {
      pos++
      const inner = expr()
      if (inner === null) return null
      const closing = peek()
      if (!closing || closing.kind !== 'op' || closing.value !== ')') return null
      pos++
      return inner
    }
    return null
  }
  const result = expr()
  if (result === null || pos !== tokens.length) return null // trailing garbage -> undecidable
  return Number.isFinite(result) ? result : null // Infinity/NaN (e.g. x/0) -> undecidable
}

/**
 * Evaluate a simple four-arithmetic spreadsheet formula exactly, or declare it undecidable.
 *
 * @param {string} formula Formula text with or without the leading '=' (e.g. '=A1*B1').
 * @param {(ref: string) => number|null|undefined} resolveCell Resolves a normalized reference
 *        ('$' stripped, uppercased, e.g. 'A1') to the replayed numeric value; a missing or
 *        non-numeric reference must come back as null/undefined/non-number and makes the result
 *        null (the caller records 人工核对).
 * @returns {number|null} The exact double result, or null when the formula is outside the
 *          whitelist (functions/strings/ranges/unknown tokens), references are missing or
 *          non-numeric, or the value is not finite.
 */
function evalSimpleFormula(formula, resolveCell) {
  if (typeof formula !== 'string' || typeof resolveCell !== 'function') return null
  let text = formula.trim()
  if (text.startsWith('=')) text = text.slice(1).trim()
  if (!text) return null
  // Whitelist scan first (defense in depth): a formula that contains ANY character outside the
  // simple grammar is rejected before parsing, so the recursive descent below only ever sees
  // tokens it understands.
  if (!SIMPLE_FORMULA.test(text.replace(/\s+/g, ''))) return null
  return evaluate(text, resolveCell)
}

export { evalSimpleFormula }

// Parse a Qty cell string. If the user typed a strict
// "<number> [x|X|*] <number>" expression, return the product as a
// string. Otherwise return the input unchanged so plain numbers and
// invalid inputs flow through to the existing PATCH normalizer.
//
// The regex is intentionally narrow — it requires exactly two numeric
// operands and a single x/X/* between them — so values like
// "2x4 SF" or "12 x 8 + extra" never silently multiply.

const MUL_RE = /^\s*(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)\s*$/

export function parseQty(raw) {
  if (raw == null || raw === '') return raw
  const m = MUL_RE.exec(String(raw))
  if (!m) return raw
  const product = parseFloat(m[1]) * parseFloat(m[2])
  if (!isFinite(product)) return raw
  return String(product)
}

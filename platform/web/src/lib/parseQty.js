// Parse a Qty cell string. Accepts plain numbers and short
// arithmetic expressions made of `+`, `*` (or `x` / `X`), and
// numeric literals — `*` binds tighter than `+`, the conventional
// math precedence.
//
// Examples:
//   "15"          → "15"
//   "15x10"       → "150"
//   "15*10"       → "150"
//   "15+10"       → "25"
//   "15+10+20"    → "45"            (cumulative addition for block-by-block measurements)
//   "15*10+12"    → "162"           (run length × width + a tail)
//   "15*2+10*3"   → "60"            (mixed)
//   "15.5+10.5"   → "26"            (decimals)
//   "abc"         → "abc"           (no match — falls through to the existing PATCH normalizer)
//   "15+"         → "15+"           (incomplete — passthrough)
//
// The expression must be a closed chain: every operator sits between
// two valid numeric operands.  Anything else (units, stray text,
// dangling operators) is returned untouched so the existing PATCH
// normalizer can do its thing.

const EXPR_RE = /^\s*(\d+(?:\.\d+)?)(\s*[+xX*]\s*\d+(?:\.\d+)?)*\s*$/

export function parseQty(raw) {
  if (raw == null || raw === '') return raw
  const s = String(raw)
  if (!EXPR_RE.test(s)) return raw

  // Split on '+' for addition terms, then each term on '*'/'x'/'X'
  // for multiplication factors.  This naturally enforces the
  // standard precedence (multiply within a term, then sum terms).
  let total = 0
  for (const term of s.split('+')) {
    const factors = term.split(/[*xX]/).map(f => parseFloat(f.trim()))
    if (factors.some(f => !isFinite(f))) return raw
    const product = factors.reduce((a, b) => a * b, 1)
    total += product
  }
  if (!isFinite(total)) return raw
  return String(total)
}

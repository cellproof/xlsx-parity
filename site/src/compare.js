// Same rule as suite/run.py excel_match / EXCEL_REL_TOL.
// Bool is not a number. Strings and errors are exact. Numbers use
// math.isclose(rel_tol=1e-12, abs_tol=0): exact equality first, then
// abs(a-b) <= rel_tol * abs(a) or abs(a-b) <= rel_tol * abs(b).

export const EXCEL_REL_TOL = 1e-12;

export function isClose(a, b, relTol = EXCEL_REL_TOL) {
  if (a === b) return true;
  if (!Number.isFinite(a) || !Number.isFinite(b)) return false;
  const diff = Math.abs(a - b);
  return diff <= Math.abs(relTol * b) || diff <= Math.abs(relTol * a);
}

export function excelMatch(got, expected) {
  const gotBool = typeof got === "boolean";
  const expBool = typeof expected === "boolean";
  if (gotBool || expBool) return gotBool && expBool && got === expected;
  if (typeof expected === "string" || typeof got === "string") return got === expected;
  if (typeof expected === "number" && typeof got === "number") return isClose(got, expected);
  return false;
}

export function valuesEqual(a, b) {
  const aBool = typeof a === "boolean";
  const bBool = typeof b === "boolean";
  if (aBool || bBool) return aBool && bBool && a === b;
  if (typeof a === "string" || typeof b === "string") return a === b;
  if (typeof a === "number" && typeof b === "number") return a === b;
  return a === b;
}

export function isErrorString(value) {
  return typeof value === "string" && ((value.startsWith("#") && (value.endsWith("!") || value.endsWith("?"))) || value === "#N/A");
}

// suite/run.py show(): errors stay bare, strings keep JSON quotes.
export function showValue(value) {
  if (typeof value === "string") {
    if (isErrorString(value)) return value;
    return JSON.stringify(value);
  }
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "number") {
    if (Object.is(value, -0)) return "-0";
    if (Number.isInteger(value) && Math.abs(value) <= Number.MAX_SAFE_INTEGER) return String(value);
    return String(value);
  }
  if (value == null) return "";
  return String(value);
}

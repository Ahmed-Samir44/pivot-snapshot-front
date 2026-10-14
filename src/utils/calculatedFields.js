// Calculated fields (Excel's Insert Calculated Field) — since 2026-10-14 written as Excel's own free
// formula ("= 'Total Revenue' / 'Net Volume'"), which the backend parses safely
// (CalculatedFieldFormula.cs). A field saved before that has leftField/operator/rightField instead.

const OPERATOR_SYMBOLS = { Add: "+", Subtract: "-", Multiply: "*", Divide: "/" };

// A field name as Excel writes it in a formula: bare when it's one plain word, else 'quoted'
// ('' for a literal quote).
export function formulaFieldReference(name) {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

// The formula a calculated field shows in the dialog — its own, or one built from the older shape.
export function calculatedFieldFormula(field) {
  if (field.formula) return field.formula;
  if (field.leftField && field.rightField) {
    return `= ${formulaFieldReference(field.leftField)} ${OPERATOR_SYMBOLS[field.operator] ?? "/"} ${formulaFieldReference(field.rightField)}`;
  }
  return "= 0";
}

// Complete enough to send: a name plus a formula (or both measures of the older shape).
export function isCompleteCalculatedField(field) {
  return Boolean(field.name) && (Boolean(field.formula?.trim()) || (Boolean(field.leftField) && Boolean(field.rightField)));
}

// Calculated items (Excel's Insert Calculated Item) — the same formula style, where a name is an ITEM
// of the item's own field. One saved before 2026-10-14 has positiveMembers/negativeMembers instead.
export function calculatedItemFormula(item) {
  if (item.formula) return item.formula;
  const plus = (item.positiveMembers ?? []).map(formulaFieldReference);
  const minus = (item.negativeMembers ?? []).map((m) => `- ${formulaFieldReference(m)}`);
  if (plus.length + minus.length === 0) return "= 0";
  return `= ${[plus.join(" + "), ...minus].filter(Boolean).join(" ")}`;
}

export function isCompleteCalculatedItem(item) {
  return (
    Boolean(item.field) &&
    Boolean(item.name) &&
    (Boolean(item.formula?.trim()) || (item.positiveMembers ?? []).length > 0 || (item.negativeMembers ?? []).length > 0)
  );
}

// Generated once from a run of the compare, then kept by hand: the cells of the
// IronCalc calc_tests workbooks that this app does not yet work out as Excel stored them,
// by workbook, with how they differ:
//   legacy        this app spills (or #SPILL! stops it) where Excel gives one value (none left since the @ on import, 2026-11-19)
//   excel-error   Excel gives an error (#NUM!, #VALUE! ...) where this app gives a value
//   we-error      this app gives an error where Excel gives a value
//   value         both give a value, and they differ
// A cell leaves this list when it comes out right (the test says so).
export const KNOWN: Record<string, Record<string, string>> = {
  "DATE_AND_TIME/DATEVALUE.xlsx": {
    "DATEVALUE!B48": "we-error",
  },
  "DATE_AND_TIME/DAYS_DAYS360.xlsx": {
    "Sheet1!C71": "excel-error",
    "Sheet1!D71": "excel-error",
  },
  "DATE_AND_TIME/NETWORKDAYS_NETWORKDAYS.INTL.xlsx": {
    "Sheet1!C60": "excel-error",
    "Sheet1!D60": "excel-error",
  },
  "FINANCIAL/ACCRINT.xlsx": {
    "Sheet1!B5": "value",
    "Sheet1!B32": "value",
    "Sheet1!B33": "value",
    "Sheet1!B34": "value",
    "Sheet1!B36": "value",
    "Sheet1!B38": "value",
    "Sheet1!B39": "value",
  },
  "FINANCIAL/RATE.xlsx": {
    "Sheet1!F17": "excel-error",
    "Sheet1!G17": "excel-error",
  },
  "FINANCIAL/XNPV.xlsx": {
    "Sheet1!R2": "value",
    "Sheet1!R5": "value",
  },
  "STATISTICAL/LOGNORM.xlsx": {
    "LOGNORM.DIST LOGNORM.INV!F26": "value",
  },
  "TEXT/T_VALUE_VALUETOTEXT.xlsx": {
    "Sheet1!I5": "value",
    "Sheet1!C15": "value",
  },
  "MATH_AND_TRIGONOMETRY/TRIGONOMETRIC.xlsx": {
    "SIN_COS_TAN!D45": "value",
    "SIN_COS_TAN!D46": "value",
    "COT_CSC_SEC!B44": "value",
    "COT_CSC_SEC!C44": "value",
    "COT_CSC_SEC!D45": "value",
    "COT_CSC_SEC!D46": "value",
  },
  "spill_operator.xlsx": {
    "Sheet1!G2": "we-error",
  },
};

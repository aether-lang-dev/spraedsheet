// Generated once from a run of the compare, then kept by hand: the cells of the
// IronCalc calc_tests workbooks that this app does not yet work out as Excel stored them,
// by workbook, with how they differ:
//   legacy        the formula has no array flag, so Excel intersects a range implicitly; this app spills
//   excel-error   Excel gives an error (#NUM!, #VALUE! ...) where this app gives a value
//   we-error      this app gives an error where Excel gives a value
//   value         both give a value, and they differ
// A cell leaves this list when it comes out right (the test says so).
export const KNOWN: Record<string, Record<string, string>> = {
  "DATE_AND_TIME/TIMEVALUE.xlsx": {
    "TIMEVALUE!A22": "legacy",
    "TIMEVALUE!B22": "legacy",
  },
  "DATE_AND_TIME/YEARFRAC_DATEDIF.xlsx": {
    "YEARFRAC_DATEDIF!A30": "legacy",
    "YEARFRAC_DATEDIF!D30": "legacy",
    "YEARFRAC_DATEDIF!B31": "legacy",
    "YEARFRAC_DATEDIF!D31": "legacy",
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
  "INFORMATION/ERROR.TYPE.xlsx": {
    "Sheet1!B8": "value",
    "Sheet1!B14": "value",
  },
  "INFORMATION/ISREF.xlsx": {
    "Sheet1!D2": "excel-error",
    "Sheet1!D7": "excel-error",
    "Sheet1!D9": "legacy",
  },
  "LOOKUP_AND_REFERENCE/CHOOSE.xlsx": {
    "CHOOSE!B23": "legacy",
  },
  "LOOKUP_AND_REFERENCE/FILTER_ERRORS.xlsx": {
    "Sheet1!E1": "excel-error",
  },
  "MATH_AND_TRIGONOMETRY/ARABIC_ROMAN.xlsx": {
    "ROMAN!A36": "legacy",
    "ROMAN!B36": "legacy",
    "ROMAN!C36": "legacy",
    "ROMAN!D36": "legacy",
    "ROMAN!E36": "legacy",
    "ROMAN!F36": "legacy",
    "ARABIC!A41": "legacy",
    "ARABIC!B41": "legacy",
  },
  "MATH_AND_TRIGONOMETRY/COMBIN_COMBINA.xlsx": {
    "COMBIN_COMBINA!C24": "legacy",
  },
  "MATH_AND_TRIGONOMETRY/DEGREES_RADIANS.xlsx": {
    "DEGREES_RADIANS!A18": "legacy",
    "DEGREES_RADIANS!B18": "legacy",
    "DEGREES_RADIANS!A42": "legacy",
    "DEGREES_RADIANS!B42": "legacy",
  },
  "MATH_AND_TRIGONOMETRY/EVEN_ODD.xlsx": {
    "EVEN_ODD!A39": "legacy",
    "EVEN_ODD!B39": "legacy",
    "EVEN_ODD!C39": "legacy",
  },
  "MATH_AND_TRIGONOMETRY/EXP_SIGN.xlsx": {
    "EXP_SIGN!B36": "legacy",
    "EXP_SIGN!C36": "legacy",
  },
  "MATH_AND_TRIGONOMETRY/MOD_QUOTIENT.xlsx": {
    "DATEVALUE!A41": "legacy",
    "DATEVALUE!C41": "legacy",
    "DATEVALUE!D41": "legacy",
  },
  "STATISTICAL/EXPONDIST.xlsx": {
    "EXPON.DIST!D56": "we-error",
  },
  "STATISTICAL/FISHER_FISHERINV.xlsx": {
    "FISHER FISHERINV!B51": "value",
    "FISHER FISHERINV!C51": "value",
    "FISHER FISHERINV!C52": "value",
  },
  "STATISTICAL/GAMMALN_GAMMALNPRECISE.xlsx": {
    "GAMMALN GAMMALN.PRECISE!B45": "we-error",
    "GAMMALN GAMMALN.PRECISE!C45": "we-error",
  },
  "STATISTICAL/LOGNORM.xlsx": {
    "LOGNORM.DIST LOGNORM.INV!F26": "value",
    "LOGNORM.DIST LOGNORM.INV!E66": "legacy",
    "LOGNORM.DIST LOGNORM.INV!F66": "legacy",
    "LOGNORM.DIST LOGNORM.INV!E67": "legacy",
    "LOGNORM.DIST LOGNORM.INV!F67": "legacy",
    "LOGNORM.DIST LOGNORM.INV!E68": "we-error",
    "LOGNORM.DIST LOGNORM.INV!F68": "we-error",
  },
  "STATISTICAL/NORM_DIST.xlsx": {
    "NORM.DIST NORM.INV!E64": "value",
    "NORM.DIST NORM.INV!F64": "legacy",
    "NORM.S.DIST NORM.S.INV!C48": "legacy",
    "NORM.S.DIST NORM.S.INV!D48": "legacy",
    "NORM.S.DIST NORM.S.INV!C49": "value",
  },
  "STATISTICAL/POISSON.xlsx": {
    "POISSON.DIST!D60": "legacy",
    "POISSON.DIST!D61": "value",
  },
  "STATISTICAL/T_DIST.xlsx": {
    "T.DIST T.INV T.DIST.RT!E19": "value",
    "T.DIST.2T T.INV.2T!D62": "value",
  },
  "TEXT/CONCATENATE.xlsx": {
    "Sheet2!B2": "legacy",
  },
  "TEXT/T_VALUE_VALUETOTEXT.xlsx": {
    "Sheet1!I5": "value",
    "Sheet1!C15": "value",
  },
  "spill_operator.xlsx": {
    "Sheet1!G2": "we-error",
    "Sheet1!I2": "we-error",
  },
};

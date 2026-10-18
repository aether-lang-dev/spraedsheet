// Number-format cases from IronCalc (MIT; third_party/ironcalc/PROVENANCE.md),
// run against app/sheet/numfmt.ts as data. Each row is
//   [upstream test, value, format, text IronCalc expects, text we expect]
// and the last is left out when the two agree. Where they differ, Excel's
// documented behaviour decided which is right (the reason in the row's note),
// not IronCalc's value. From base/src/test/test_number_format.rs and
// base/src/formatter/test/{test_en_examples,test_general,test_scientific,
// test_time,test_dates}.rs (English locale; the de/fr files are the locale
// work, not ours yet).
import { numberFormat } from "./sheet/numfmt.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

type Row = [string, Value, string, string | null] | [string, Value, string, string | null, string | null, string];
const ROWS: Row[] = [
  // rounding and decimals
  ["test_en_examples::format_0", 5.6, "0", "6"], ["format_0", 0.2, "0", "0"], ["format_0", 123, "0", "123"],
  ["format_0_00", 5.6, "0.00", "5.60"], ["format_0_00", 0.2, "0.00", "0.20"], ["format_0_00", 123.456, "0.00", "123.46"],
  ["format_hash_group_0", 1234, "#,##0", "1,234"], ["format_hash_group_0", 12, "#,##0", "12"], ["format_hash_group_0", 0.4, "#,##0", "0"],
  ["format_hash_group_0_00", 1234.5, "#,##0.00", "1,234.50"], ["format_hash_group_0_00", 0.456, "#,##0.00", "0.46"],
  ["format_hash_group_0_hashhashhash", 1234.5678, "#,##0.###", "1,234.568"], ["format_hash_group_0_hashhashhash", 12.3, "#,##0.###", "12.3"],
  ["format_hash_group_0_hashhashhash", 0.5, "#,##0.###", "0.5"],
  ["format_hash_only", 0, "#", ""], ["format_hash_only", 12, "#", "12"], ["format_hash_only", 0.6, "#", "1"],
  ["format_aaa0", 7, "###0", "7"], ["format_aaa0", 0.4, "###0", "0"], ["format_aaa0", 1234, "###0", "1234"],
  ["format_0_hashhashhash", 1.2346, "0.###", "1.235"], ["format_0_hashhashhash", 1.2, "0.###", "1.2"], ["format_0_hashhashhash", 0.05, "0.###", "0.05"],
  ["format_hash_group_0_0hash", 1234.567, "#,##0.0#", "1,234.57"], ["format_hash_group_0_0hash", 1234.5, "#,##0.0#", "1,234.5"], ["format_hash_group_0_0hash", 12, "#,##0.0#", "12.0"],
  ["format_group_with_leading_hashes (ignored upstream)", 1234.5, "###,##0.00", "1,234.50"], ["format_group_with_leading_hashes", 12, "###,##0.00", "12.00"], ["format_group_with_leading_hashes", 0.2, "###,##0.00", "0.20"],
  ["test_number_format::test_simple_format", 2.3, "General", "2.3"],
  ["test_general::test_decimal_part", 3.1, "0.00", "3.10"],
  // Excel shows 15 significant digits and zeros after them
  ["test_number_format::test_maximum_zeros", 1 / 3, "#,##0.0000000000000000000", "0.3333333333333330000"],
  ["test_maximum_zeros", 1234 + 1 / 3, "#,##0.0000000000000000000", "1,234.3333333333300000000"],
  ["test_general::test_percent_correct_rounding", 0.1399999999999999, "0.0%", "14.0%"], ["test_percent_correct_rounding", -0.1399999999999999, "0.0%", "-14.0%"],
  ["test_percent_correct_rounding", 0.1399, "0.0%", "14.0%"], ["test_percent_correct_rounding", 0.02666, "0.00%", "2.67%"], ["test_percent_correct_rounding", 0.0266, "0.00%", "2.66%"],
  ["test_percent_correct_rounding", 0.0233, "0.00%", "2.33%"], ["test_percent_correct_rounding", 0.02666, "0%", "3%"], ["test_percent_correct_rounding", -0.02666, "0.00%", "-2.67%"],
  ["test_percent_correct_rounding", -0.02666, "0%", "-3%"], ["test_percent_correct_rounding", 0.135, "0%", "14%"], ["test_percent_correct_rounding", 0.13499, "0%", "13%"],
  ["test_percent_correct_rounding", -0.135, "0%", "-14%"], ["test_percent_correct_rounding", -0.13499, "0%", "-13%"], ["test_percent_correct_rounding", 0.1345, "0.0%", "13.5%"],
  ["test_percent_correct_rounding", 0.1343, "0.0%", "13.4%"], ["test_percent_correct_rounding", -0.1345, "0.0%", "-13.5%"], ["test_percent_correct_rounding", -0.134499, "0.0%", "-13.4%"],
  // percent
  ["format_0_percent", 0.256, "0%", "26%"], ["format_0_percent", 1, "0%", "100%"], ["format_0_percent", 0.004, "0%", "0%"],
  ["format_0_00_percent", 0.256, "0.00%", "25.60%"], ["format_0_00_percent", 1, "0.00%", "100.00%"], ["format_0_00_percent", 0.004, "0.00%", "0.40%"],
  ["format_0_hashhash_percent", 0.256, "0.##%", "25.6%"], ["format_0_hashhash_percent", 1, "0.##%", "100.%"], ["format_0_hashhash_percent", 0.0043, "0.##%", "0.43%"],
  ["test_percent", 0.12, "0.00%", "12.00%"], ["test_percent", 0.12, "0.00%%", "1200.00%%"],
  // thousands separators and scaling
  ["format_thousands_scale", 1234567, "#,##0,", "1,235"], ["format_thousands_scale", 1500, "#,##0,", "2"], ["format_thousands_scale", 999, "#,##0,", "1"],
  ["format_millions_scale_integer", 1234567, "#,##0,,", "1"], ["format_millions_scale_integer", 2510000, "#,##0,,", "3"], ["format_millions_scale_integer", 12345678, "#,##0,,", "12"],
  ["format_millions_scale_decimal", 1234567, "#,##0,,.0#", "1.23"], ["format_millions_scale_decimal", 12000000, "#,##0,,.0#", "12.0"], ["format_millions_scale_decimal", 987654321, "#,##0,,.0#", "987.65"],
  ["test_general::simple_test", 7, "000", "007"], ["simple_test_comma", 1007, "000", "1007"], ["simple_test_comma", 1008, "#", "1008"], ["simple_test_comma", 1009, "#,#", "1,009"],
  ["simple_test_comma", 12345678, "#,#", "12,345,678"], ["simple_test_comma", 12345678, "0,0", "12,345,678"], ["simple_test_comma", 1005, "00-00", "10-05"],
  ["simple_test_comma", 7, "0?0", "0 7"], ["simple_test_comma", 7, "0#0", "07"], ["simple_test_comma", 1234, '000 "Millions"', "1234 Millions"],
  ["simple_test_comma", 1235, '#,000 "Millions"', "1,235 Millions"], ["simple_test_comma", 1007, "0,00", "1,007"], ["simple_test_comma", 10000007, "0,00", "10,000,007"],
  // sections and signs
  ["test_negative_numbers", -123, "0.0", "-123.0"], ["test_negative_numbers", -3, "000.0", "-003.0"],
  ["test_negative_numbers", -0.00001, "000.0", "000.0", "000.0", "kept: Excel shows -000.0 here, but our arithmetic has no near-zero snap (0.3-0.1-0.2 would show -0.00 where Excel shows 0.00), so a negative that rounds to zero shows no sign"],
  ["test_zero", 0, "$#,##0", "$0"], ["test_zero", -1 / 3, "0", "0", "0", "kept: see test_negative_numbers"], ["test_zero", -1 / 3, "0;(0)", "(0)"],
  ["test_parts", 3.1, "0.00;(0.00);(-)", "3.10"], ["test_parts", -3.1, "0.00;(0.00);(-)", "(3.10)"], ["test_parts", 0, "0.00;(0.00);(-)", "(-)"],
  ["test_negative_currencies", -23, "$#,##0", "-$23"], ["test_currency", 123.1, "$#,##0", "$123"], ["test_currency", 123.1, "#,##0 €", "123 €"],
  ["test_currency_valid_formats", 2.34, "[$€]#,##0.00", "€2.34"], ["test_currency_valid_formats", 3, "[$€]#,##0.00", "€3.00"], ["test_currency_valid_formats", 42, "[$€]0", "€42"],
  ["test_conditional", 42, "[=]0.00", "42.00"],
  ["test_string_double_quota", 42, '0 """Millions"""', '42 "Millions"', "42 Millions", "kept: in a number format a quote is written \\\" (backslash); a doubled quote is the empty string, as Excel reads it. IronCalc reads it the way a formula's string does"],
  // scientific
  ["test_scientific::scientific_minus_negative", 0.000002, "0.00E-00", "2.00E-06"], ["scientific_minus_positive", 2000000, "0.00E-00", "2.00E06"],
  ["scientific_positive", 2000000, "0.00E+00", "2.00E+06"], ["scientific_negative", 0.000002, "0.00E+00", "2.00E-06"],
  ["test_general::test_scientific", 2.5e-14, "0.00E+0", "2.50E-14"], ["test_scientific", 3e-4, "0.00E+00", "3.00E-04"],
  ["format_scientific", 12345.1, "0.###E+00", "1.235E+04"], ["format_scientific", 0.0123, "0.###E+00", "1.23E-02"], ["format_scientific", 1000, "0.###E+00", "1.E+03"],
  ["format_scientific_failing (ignored upstream: wrong rounding)", 12345, "0.###E+00", "1.235E+04"],
  // fractions (ignored upstream: not implemented there)
  ["format_fraction_simple", 1.25, "# ?/?", "1 1/4"], ["format_fraction_simple", 2.5, "# ?/?", "2 1/2"], ["format_fraction_simple", 0.5, "# ?/?", " 1/2"],
  ["format_fraction_two_digit_den", 1.333, "# ?/??", "1 1/3", "1 1/3 ", "kept: ? holds its place with a space, so the unused ? of ?/?? pads (Excel's own documentation of ?)"], ["format_fraction_two_digit_den", 1.2, "# ?/??", "1 1/5", "1 1/5 ", "kept: as above"], ["format_fraction_two_digit_den", 0.75, "# ?/??", " 3/4", " 3/4 ", "kept: as above"],
  ["format_fraction_eighths", 1.25, "# ?/8", "1 2/8"], ["format_fraction_eighths", 0.5, "# ?/8", " 4/8"], ["format_fraction_eighths", 2.75, "# ?/8", "2 6/8"],
  // dates and times
  ["test_dates::simple_test", 46015, "m.d.yy", "12.24.25"],
  ["test_general::test_date", 41181, "dd/mm/yyyy", "29/09/2012"], ["test_date", 41181, "dd-mm-yyyy", "29-09-2012"], ["test_date", 41304, "dd-mm-yyyy", "30-01-2013"],
  ["test_date", 42657, "dd-mm-yyyy", "14-10-2016"], ["test_date", 41181, "dddd-mmmm-yyyy", "Saturday-September-2012"], ["test_date", 41181, "ddd-mmm-yy", "Sat-Sep-12"],
  ["test_date", 41181, "ddd-mmmmm-yy", "Sat-S-12"],
  ["test_date", 41181, "ddd-mmmmmmm-yy", "Sat-September-12"],
  ["test_time::simple_test", 16.00142361111111, "h:mm AM/PM", "12:02 AM"],
  ["padded_vs_unpadded", 0.25351851851851853, "hh:mm:ss AM/PM", "06:05:04 AM"], ["padded_vs_unpadded", 0.25351851851851853, "h:m:s AM/PM", "6:5:4 AM"],
  ["elapsed_hour", 1.5, "[hh]:mm:ss", "36:00:00"], ["elapsed_minute", 2, "[mm]:ss", "2880:00"], ["elapsed_second", 0.5, "[ss]", "43200"],
  ["elapsed_hour_padded", 0.1, "[hh]:mm:ss", "02:24:00"],
  // text is left as it is
  ["test_leading_comma_text", ",10", "General", null], ["test_leading_comma_text", ",1000", "General", null],
];

const diffs: string[] = [];
const notes: string[] = [];
for (const row of ROWS) {
  const [name, v, pattern, upstream] = row;
  const want = row.length === 6 ? row[4] : upstream;
  const got = numberFormat(v, pattern)?.text ?? null;
  if (got !== want) diffs.push(`${name} ${JSON.stringify(v)} ${pattern}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
  if (row.length === 6) notes.push(name);
}
eq("ironcalc number formats: the upstream cases", diffs, []);
eq("ironcalc number formats: the cases where we deliberately differ from IronCalc", notes, ["test_negative_numbers", "test_zero", "test_string_double_quota", "format_fraction_two_digit_den", "format_fraction_two_digit_den", "format_fraction_two_digit_den"]);
// expect: ok ironcalc number formats: the upstream cases
// expect: ok ironcalc number formats: the cases where we deliberately differ from IronCalc

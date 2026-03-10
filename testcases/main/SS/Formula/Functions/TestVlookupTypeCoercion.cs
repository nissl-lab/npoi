/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests that VLOOKUP correctly coerces types when comparing strings and numbers,
    /// matching Excel's actual behavior.
    ///
    /// <b>Background:</b> In Excel, VLOOKUP/HLOOKUP/MATCH coerce types during comparison.
    /// For example, a cell containing the text "10" will match a lookup table key containing
    /// the number 10, and vice versa. This is easily verified by formatting a cell as Text,
    /// entering "10", and running VLOOKUP against a column of numbers — Excel finds the match.
    ///
    /// NPOI's original implementation returned TypeMismatch for all cross-type comparisons,
    /// causing lookups to produce #N/A in situations where Excel would succeed. Since the
    /// purpose of NPOI's formula evaluator is to replicate Excel's behavior, this fix aligns
    /// NPOI with Excel's actual semantics.
    ///
    /// This is a behavior change from previous NPOI versions. Code that relied on the strict
    /// type separation (string "10" NOT matching number 10) was depending on behavior that
    /// diverged from Excel. Such code would already produce different results when the same
    /// spreadsheet was opened in Excel.
    /// </summary>
    [TestFixture]
    public class TestVlookupTypeCoercion
    {
        /// <summary>
        /// Tests VLOOKUP where the lookup value is a string but the table keys are numbers.
        /// Example: cell contains string "25000" but lookup table keys are numeric.
        /// </summary>
        [Test]
        public void TestStringLookupValueMatchesNumericTableKey()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Lookup table in A1:B3 with numeric keys
            IRow row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue(10000); // number
            row1.CreateCell(1).SetCellValue(1.5);   // factor

            IRow row2 = sheet.CreateRow(1);
            row2.CreateCell(0).SetCellValue(25000); // number
            row2.CreateCell(1).SetCellValue(2.0);

            IRow row3 = sheet.CreateRow(2);
            row3.CreateCell(0).SetCellValue(50000); // number
            row3.CreateCell(1).SetCellValue(3.0);

            // Lookup value in C1 as STRING "25000"
            ICell lookupCell = row1.CreateCell(2);
            lookupCell.SetCellValue("25000"); // string, not number

            // VLOOKUP formula in D1
            ICell formulaCell = row1.CreateCell(3);
            formulaCell.SetCellFormula("VLOOKUP(C1,A1:B3,2,FALSE)");

            CellValue result = evaluator.Evaluate(formulaCell);
            ClassicAssert.AreEqual(CellType.Numeric, result.CellType,
                "String lookup '25000' should match numeric key 25000");
            ClassicAssert.AreEqual(2.0, result.NumberValue, 0.0001);
        }

        /// <summary>
        /// Tests VLOOKUP where the lookup value is a number but the table keys are strings.
        /// Example: spreadsheet stores codes as text ("1", "2", "10") but the
        /// lookup cell contains a number.
        /// </summary>
        [Test]
        public void TestNumericLookupValueMatchesStringTableKey()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Lookup table in A1:B3 with string keys
            IRow row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue("1");   // string
            row1.CreateCell(1).SetCellValue(100);

            IRow row2 = sheet.CreateRow(1);
            row2.CreateCell(0).SetCellValue("5");   // string
            row2.CreateCell(1).SetCellValue(200);

            IRow row3 = sheet.CreateRow(2);
            row3.CreateCell(0).SetCellValue("10");  // string
            row3.CreateCell(1).SetCellValue(300);

            // Lookup value in C1 as NUMBER 10
            ICell lookupCell = row1.CreateCell(2);
            lookupCell.SetCellValue(10); // number, not string

            // VLOOKUP formula in D1
            ICell formulaCell = row1.CreateCell(3);
            formulaCell.SetCellFormula("VLOOKUP(C1,A1:B3,2,FALSE)");

            CellValue result = evaluator.Evaluate(formulaCell);
            ClassicAssert.AreEqual(CellType.Numeric, result.CellType,
                "Numeric lookup 10 should match string key '10'");
            ClassicAssert.AreEqual(300.0, result.NumberValue, 0.0001);
        }

        /// <summary>
        /// Tests VLOOKUP with decimal values — string "0.02" should match number 0.02.
        /// </summary>
        [Test]
        public void TestStringDecimalLookupMatchesNumericKey()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            IRow row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue(0.01);  // number
            row1.CreateCell(1).SetCellValue("Low");

            IRow row2 = sheet.CreateRow(1);
            row2.CreateCell(0).SetCellValue(0.02);  // number
            row2.CreateCell(1).SetCellValue("Medium");

            IRow row3 = sheet.CreateRow(2);
            row3.CreateCell(0).SetCellValue(0.05);  // number
            row3.CreateCell(1).SetCellValue("High");

            // Lookup value as STRING "0.02"
            ICell lookupCell = row1.CreateCell(2);
            lookupCell.SetCellValue("0.02");

            ICell formulaCell = row1.CreateCell(3);
            formulaCell.SetCellFormula("VLOOKUP(C1,A1:B3,2,FALSE)");

            CellValue result = evaluator.Evaluate(formulaCell);
            ClassicAssert.AreEqual(CellType.String, result.CellType,
                "String '0.02' should match numeric key 0.02");
            ClassicAssert.AreEqual("Medium", result.StringValue);
        }

        /// <summary>
        /// Tests that non-numeric strings still produce #N/A when the table has numeric keys.
        /// Type coercion should only work when the string is actually parseable as a number.
        /// </summary>
        [Test]
        public void TestNonNumericStringDoesNotMatchNumericKey()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            IRow row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue(10);
            row1.CreateCell(1).SetCellValue("found");

            // Lookup value is a non-numeric string
            ICell lookupCell = row1.CreateCell(2);
            lookupCell.SetCellValue("abc");

            ICell formulaCell = row1.CreateCell(3);
            formulaCell.SetCellFormula("VLOOKUP(C1,A1:B1,2,FALSE)");

            CellValue result = evaluator.Evaluate(formulaCell);
            ClassicAssert.AreEqual(CellType.Error, result.CellType,
                "Non-numeric string should not match numeric key");
            ClassicAssert.AreEqual(FormulaError.NA.Code, result.ErrorValue,
                "Should produce #N/A error");
        }

        /// <summary>
        /// Tests that same-type comparisons still work correctly after the coercion change.
        /// String-to-string and number-to-number should be unaffected.
        /// </summary>
        [Test]
        public void TestSameTypeComparisonsUnaffected()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Number-to-number lookup
            IRow row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue(42);
            row1.CreateCell(1).SetCellValue("answer");

            ICell lookupCell1 = row1.CreateCell(2);
            lookupCell1.SetCellValue(42);

            ICell formulaCell1 = row1.CreateCell(3);
            formulaCell1.SetCellFormula("VLOOKUP(C1,A1:B1,2,FALSE)");

            CellValue result1 = evaluator.Evaluate(formulaCell1);
            ClassicAssert.AreEqual("answer", result1.StringValue,
                "Number-to-number VLOOKUP should still work");

            // String-to-string lookup
            IRow row2 = sheet.CreateRow(1);
            row2.CreateCell(0).SetCellValue("key");
            row2.CreateCell(1).SetCellValue("value");

            ICell lookupCell2 = row2.CreateCell(2);
            lookupCell2.SetCellValue("key");

            ICell formulaCell2 = row2.CreateCell(3);
            formulaCell2.SetCellFormula("VLOOKUP(C2,A2:B2,2,FALSE)");

            CellValue result2 = evaluator.Evaluate(formulaCell2);
            ClassicAssert.AreEqual("value", result2.StringValue,
                "String-to-string VLOOKUP should still work");
        }

        /// <summary>
        /// Tests VLOOKUP with approximate match (sorted data) and type coercion.
        /// Ensures coercion works with range_lookup=TRUE as well.
        /// </summary>
        [Test]
        public void TestApproximateMatchWithTypeCoercion()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Sorted numeric table
            IRow row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue(10);
            row1.CreateCell(1).SetCellValue("ten");

            IRow row2 = sheet.CreateRow(1);
            row2.CreateCell(0).SetCellValue(20);
            row2.CreateCell(1).SetCellValue("twenty");

            IRow row3 = sheet.CreateRow(2);
            row3.CreateCell(0).SetCellValue(30);
            row3.CreateCell(1).SetCellValue("thirty");

            // Lookup with string "25" — approximate match should find 20 (largest <= 25)
            ICell lookupCell = row1.CreateCell(2);
            lookupCell.SetCellValue("25");

            ICell formulaCell = row1.CreateCell(3);
            formulaCell.SetCellFormula("VLOOKUP(C1,A1:B3,2,TRUE)");

            CellValue result = evaluator.Evaluate(formulaCell);
            ClassicAssert.AreEqual(CellType.String, result.CellType,
                "Approximate match with string '25' should find the row for 20");
            ClassicAssert.AreEqual("twenty", result.StringValue);
        }

        /// <summary>
        /// Tests HLOOKUP with type coercion to ensure the fix applies to both
        /// VLOOKUP and HLOOKUP (they share the same LookupValueComparer logic).
        /// </summary>
        [Test]
        public void TestHlookupTypeCoercion()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Horizontal lookup table
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue(100);  // number
            headerRow.CreateCell(1).SetCellValue(200);  // number
            headerRow.CreateCell(2).SetCellValue(300);  // number

            IRow valueRow = sheet.CreateRow(1);
            valueRow.CreateCell(0).SetCellValue("A");
            valueRow.CreateCell(1).SetCellValue("B");
            valueRow.CreateCell(2).SetCellValue("C");

            // Lookup with string "200"
            ICell lookupCell = sheet.CreateRow(2).CreateCell(0);
            lookupCell.SetCellValue("200");

            ICell formulaCell = sheet.GetRow(2).CreateCell(1);
            formulaCell.SetCellFormula("HLOOKUP(A3,A1:C2,2,FALSE)");

            CellValue result = evaluator.Evaluate(formulaCell);
            ClassicAssert.AreEqual(CellType.String, result.CellType,
                "HLOOKUP should coerce string '200' to match numeric key 200");
            ClassicAssert.AreEqual("B", result.StringValue);
        }
    }
}

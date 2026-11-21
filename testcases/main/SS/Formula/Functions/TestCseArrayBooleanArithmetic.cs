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
    using NPOI.SS.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests that CSE (Ctrl+Shift+Enter) array formulas involving boolean
    /// array arithmetic evaluate correctly.
    ///
    /// The common Excel pattern:
    ///   {INDEX(result_range, MATCH(1, (criteria_range1=val1)*(criteria_range2=val2), 0))}
    ///
    /// requires that:
    ///   1. The = operator produces arrays of BoolEval (TRUE/FALSE) in array mode
    ///   2. The * operator coerces those booleans to 1/0 during element-wise multiplication
    ///   3. MATCH finds the row where the product equals 1
    ///   4. INDEX returns the corresponding value
    ///
    /// This was broken because TwoOperandNumericOperation's array evaluator used
    /// MutableValueCollector(isReferenceBoolCounted=false), which silently dropped
    /// BoolEval values from CacheAreaEval arrays.
    /// </summary>
    [TestFixture]
    public class TestCseArrayBooleanArithmetic
    {
        /// <summary>
        /// Tests the core pattern: CSE array formula with boolean multiplication
        /// for multi-criteria lookup via INDEX/MATCH.
        ///
        /// Layout:
        ///   A1:A5 = criteria column 1 ("cat", "dog", "cat", "bird", "dog")
        ///   B1:B5 = criteria column 2 ("red", "blue", "blue", "red", "red")
        ///   C1:C5 = result column (10, 20, 30, 40, 50)
        ///   D1    = lookup value 1 ("cat")
        ///   E1    = lookup value 2 ("blue")
        ///   F1    = {INDEX(C1:C5, MATCH(1, (A1:A5=D1)*(B1:B5=E1), 0))}
        ///           should return 30 (row 3: cat + blue)
        /// </summary>
        [Test]
        public void TestIndexMatchWithBooleanArrayMultiplication()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Set up criteria column 1 (A1:A5)
            string[] col1 = { "cat", "dog", "cat", "bird", "dog" };
            // Set up criteria column 2 (B1:B5)
            string[] col2 = { "red", "blue", "blue", "red", "red" };
            // Set up result column (C1:C5)
            double[] results = { 10, 20, 30, 40, 50 };

            for (int i = 0; i < 5; i++)
            {
                IRow row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue(col1[i]); // A
                row.CreateCell(1).SetCellValue(col2[i]); // B
                row.CreateCell(2).SetCellValue(results[i]); // C
            }

            // Lookup values: D1="cat", E1="blue"
            sheet.GetRow(0).CreateCell(3).SetCellValue("cat");
            sheet.GetRow(0).CreateCell(4).SetCellValue("blue");

            // CSE array formula in F1: {INDEX(C1:C5, MATCH(1, (A1:A5=D1)*(B1:B5=E1), 0))}
            sheet.SetArrayFormula(
                "INDEX(C1:C5,MATCH(1,(A1:A5=D1)*(B1:B5=E1),0))",
                CellRangeAddress.ValueOf("F1"));

            CellValue result = evaluator.Evaluate(sheet.GetRow(0).GetCell(5));

            ClassicAssert.AreEqual(CellType.Numeric, result.CellType,
                "CSE array formula with boolean multiplication should evaluate to a number");
            ClassicAssert.AreEqual(30.0, result.NumberValue, 0.0001,
                "Should find row 3 (cat+blue) and return 30");
        }

        /// <summary>
        /// Tests that boolean array multiplication works for different match positions.
        /// Verifies the result changes correctly when lookup values change.
        /// </summary>
        [Test]
        public void TestBooleanArrayMultiplicationDifferentMatches()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            string[] col1 = { "A", "B", "A", "C", "B" };
            string[] col2 = { "X", "Y", "Y", "X", "X" };
            double[] values = { 100, 200, 300, 400, 500 };

            for (int i = 0; i < 5; i++)
            {
                IRow row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue(col1[i]);
                row.CreateCell(1).SetCellValue(col2[i]);
                row.CreateCell(2).SetCellValue(values[i]);
            }

            // Test case 1: Look up B+X → should find row 5 (index 5), value 500
            sheet.GetRow(0).CreateCell(3).SetCellValue("B");
            sheet.GetRow(0).CreateCell(4).SetCellValue("X");
            sheet.SetArrayFormula(
                "INDEX(C1:C5,MATCH(1,(A1:A5=D1)*(B1:B5=E1),0))",
                CellRangeAddress.ValueOf("F1"));

            CellValue result1 = evaluator.Evaluate(sheet.GetRow(0).GetCell(5));
            ClassicAssert.AreEqual(CellType.Numeric, result1.CellType);
            ClassicAssert.AreEqual(500.0, result1.NumberValue, 0.0001,
                "B+X should match row 5, value 500");

            // Test case 2: Look up A+X → should find row 1 (index 1), value 100
            evaluator.ClearAllCachedResultValues();
            sheet.GetRow(0).GetCell(3).SetCellValue("A");
            sheet.GetRow(0).GetCell(4).SetCellValue("X");

            CellValue result2 = evaluator.Evaluate(sheet.GetRow(0).GetCell(5));
            ClassicAssert.AreEqual(CellType.Numeric, result2.CellType);
            ClassicAssert.AreEqual(100.0, result2.NumberValue, 0.0001,
                "A+X should match row 1, value 100");

            // Test case 3: Look up C+X → should find row 4 (index 4), value 400
            evaluator.ClearAllCachedResultValues();
            sheet.GetRow(0).GetCell(3).SetCellValue("C");
            sheet.GetRow(0).GetCell(4).SetCellValue("X");

            CellValue result3 = evaluator.Evaluate(sheet.GetRow(0).GetCell(5));
            ClassicAssert.AreEqual(CellType.Numeric, result3.CellType);
            ClassicAssert.AreEqual(400.0, result3.NumberValue, 0.0001,
                "C+X should match row 4, value 400");
        }

        /// <summary>
        /// Tests that when no match is found, the CSE formula returns #N/A
        /// (not #VALUE! which was the old broken behavior).
        /// </summary>
        [Test]
        public void TestBooleanArrayMultiplicationNoMatch()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            string[] col1 = { "cat", "dog" };
            string[] col2 = { "red", "blue" };
            double[] values = { 10, 20 };

            for (int i = 0; i < 2; i++)
            {
                IRow row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue(col1[i]);
                row.CreateCell(1).SetCellValue(col2[i]);
                row.CreateCell(2).SetCellValue(values[i]);
            }

            // Look up cat+blue → no match exists
            sheet.GetRow(0).CreateCell(3).SetCellValue("cat");
            sheet.GetRow(0).CreateCell(4).SetCellValue("blue");
            sheet.SetArrayFormula(
                "INDEX(C1:C2,MATCH(1,(A1:A2=D1)*(B1:B2=E1),0))",
                CellRangeAddress.ValueOf("F1"));

            CellValue result = evaluator.Evaluate(sheet.GetRow(0).GetCell(5));
            ClassicAssert.AreEqual(CellType.Error, result.CellType,
                "No match should produce an error");
            ClassicAssert.AreEqual(FormulaError.NA.Code, result.ErrorValue,
                "No match should produce #N/A (not #VALUE!)");
        }

        /// <summary>
        /// Tests that the Match function handles ErrorEval in the lookup range
        /// gracefully (returns the error) rather than throwing a raw Exception.
        /// </summary>
        [Test]
        public void TestMatchWithErrorEvalInLookupRange()
        {
            var match = new NPOI.SS.Formula.Functions.Match();
            var args = new NPOI.SS.Formula.Eval.ValueEval[]
            {
                new NPOI.SS.Formula.Eval.NumberEval(1),
                NPOI.SS.Formula.Eval.ErrorEval.VALUE_INVALID,
                new NPOI.SS.Formula.Eval.NumberEval(0)
            };

            var result = match.Evaluate(args, 0, 0);

            // Should return ErrorEval, not throw an Exception
            ClassicAssert.IsInstanceOf<NPOI.SS.Formula.Eval.ErrorEval>(result,
                "Match should return ErrorEval when lookup range is an error");
        }

        /// <summary>
        /// Tests that element-wise multiplication of two boolean arrays works
        /// correctly in CSE array formulas. This is the exact pattern used in
        /// multi-criteria lookups: (range1=val1)*(range2=val2).
        ///
        /// Layout:
        ///   A1:A4 = {1, 2, 3, 4}, B1:B4 = {1, 2, 2, 4} → (A=B) = {T, T, F, T}
        ///   C1:C4 = {10, 20, 30, 40}, D1:D4 = {10, 99, 30, 40} → (C=D) = {T, F, T, T}
        ///   CSE: {(A1:A4=B1:B4)*(C1:C4=D1:D4)} should produce {1, 0, 0, 1}
        /// </summary>
        [Test]
        public void TestSimpleBooleanArrayMultiplication()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            int[] colA = { 1, 2, 3, 4 };
            int[] colB = { 1, 2, 2, 4 };    // matches A at rows 1, 2, 4
            int[] colC = { 10, 20, 30, 40 };
            int[] colD = { 10, 99, 30, 40 }; // matches C at rows 1, 3, 4

            for (int i = 0; i < 4; i++)
            {
                IRow row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue(colA[i]);
                row.CreateCell(1).SetCellValue(colB[i]);
                row.CreateCell(2).SetCellValue(colC[i]);
                row.CreateCell(3).SetCellValue(colD[i]);
            }

            // CSE array formula: {(A1:A4=B1:B4)*(C1:C4=D1:D4)}
            // Both match at rows 1 and 4 only
            ICellRange<ICell> arrayFormula = sheet.SetArrayFormula(
                "(A1:A4=B1:B4)*(C1:C4=D1:D4)",
                CellRangeAddress.ValueOf("E1:E4"));

            // Row 1: (1=1)*(10=10) = TRUE*TRUE = 1
            ClassicAssert.AreEqual(1.0, evaluator.Evaluate(arrayFormula.FlattenedCells[0]).NumberValue, 0.0001);
            // Row 2: (2=2)*(20=99) = TRUE*FALSE = 0
            ClassicAssert.AreEqual(0.0, evaluator.Evaluate(arrayFormula.FlattenedCells[1]).NumberValue, 0.0001);
            // Row 3: (3=2)*(30=30) = FALSE*TRUE = 0
            ClassicAssert.AreEqual(0.0, evaluator.Evaluate(arrayFormula.FlattenedCells[2]).NumberValue, 0.0001);
            // Row 4: (4=4)*(40=40) = TRUE*TRUE = 1
            ClassicAssert.AreEqual(1.0, evaluator.Evaluate(arrayFormula.FlattenedCells[3]).NumberValue, 0.0001);
        }
    }
}

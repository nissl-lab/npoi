/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace TestCases.SS.Formula.Eval
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    /**
     * Tests HSSFFormulaEvaluator for its handling of cell formula circular references.
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestCircularReferences
    {
        /**
         * Translates StackOverflowError into AssertionFailedError
         */
        private static CellValue EvaluateWithCycles(HSSFWorkbook wb, ICell testCell)
        {
            HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(wb);
            try
            {
                return Evaluator.Evaluate(testCell);
            }
            catch (StackOverflowException)
            {
                throw new AssertionException("circular reference caused stack overflow error");
            }
        }
        /**
         * Makes sure that the specified Evaluated cell value represents a circular reference error.
         */
        private static void ConfirmCycleErrorCode(CellValue cellValue)
        {
            Assert.IsTrue(cellValue.CellType == CellType.ERROR);
            Assert.AreEqual(ErrorEval.CIRCULAR_REF_ERROR.ErrorCode, cellValue.ErrorValue);
        }


        /**
         * ASF Bugzilla Bug 44413
         * "INDEX() formula cannot contain its own location in the data array range"
         */
        [Test]
        public void TestIndexFormula()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");

            int colB = 1;
            sheet.CreateRow(0).CreateCell(colB).SetCellValue(1);
            sheet.CreateRow(1).CreateCell(colB).SetCellValue(2);
            sheet.CreateRow(2).CreateCell(colB).SetCellValue(3);
            IRow row4 = sheet.CreateRow(3);
            ICell testCell = row4.CreateCell(0);
            // This formula should Evaluate to the contents of B2,
            testCell.CellFormula = ("INDEX(A1:B4,2,2)");
            // However the range A1:B4 also includes the current cell A4.  If the other parameters
            // were 4 and 1, this would represent a circular reference.  Prior to v3.2 POI would
            // 'fully' Evaluate ref arguments before invoking operators, which raised the possibility of
            // cycles / StackOverflowErrors.


            CellValue cellValue = EvaluateWithCycles(wb, testCell);

            Assert.IsTrue(cellValue.CellType == CellType.NUMERIC);
            Assert.AreEqual(2, cellValue.NumberValue, 0);
        }

        /**
         * Cell A1 has formula "=A1"
         */
        [Test]
        public void TestSimpleCircularReference()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell testCell = row.CreateCell(0);
            testCell.CellFormula = ("A1");

            CellValue cellValue = EvaluateWithCycles(wb, testCell);

            ConfirmCycleErrorCode(cellValue);
        }

        /**
         * A1=B1, B1=C1, C1=D1, D1=A1
         */
        [Test]
        public void TestMultiLevelCircularReference()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).CellFormula = ("B1");
            row.CreateCell(1).CellFormula = ("C1");
            row.CreateCell(2).CellFormula = ("D1");
            ICell testCell = row.CreateCell(3);
            testCell.CellFormula = ("A1");

            CellValue cellValue = EvaluateWithCycles(wb, testCell);

            ConfirmCycleErrorCode(cellValue);
        }
        [Test]
        public void TestIntermediateCircularReferenceResults_bug46898()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");

            IRow row = sheet.CreateRow(0);

            ICell cellA1 = row.CreateCell(0);
            ICell cellB1 = row.CreateCell(1);
            ICell cellC1 = row.CreateCell(2);
            ICell cellD1 = row.CreateCell(3);
            ICell cellE1 = row.CreateCell(4);

            cellA1.CellFormula = ("IF(FALSE, 1+B1, 42)");
            cellB1.CellFormula = ("1+C1");
            cellC1.CellFormula = ("1+D1");
            cellD1.CellFormula = ("1+E1");
            cellE1.CellFormula = ("1+A1");

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv;

            // Happy day flow - Evaluate A1 first
            cv = fe.Evaluate(cellA1);
            Assert.AreEqual(CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(42.0, cv.NumberValue, 0.0);
            cv = fe.Evaluate(cellB1); // no circ-ref-error because A1 result is cached
            Assert.AreEqual(CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(46.0, cv.NumberValue, 0.0);

            // Show the bug - Evaluate another cell from the loop first
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cellB1);
            if ((int)cv.CellType == ErrorEval.CIRCULAR_REF_ERROR.ErrorCode)
            {
                throw new AssertionException("Identified bug 46898");
            }
            Assert.AreEqual(CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(46.0, cv.NumberValue, 0.0);

            // start Evaluation on another cell
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cellE1);
            Assert.AreEqual(CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(43.0, cv.NumberValue, 0.0);


        }
    }

}
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
namespace TestCases.SS.Formula.Atp
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using System;
    using TestCases.HSSF;
    using NPOI.SS.Util;
    using NPOI.SS.Formula.Eval;

    /**
     * Testcase for 'Analysis Toolpak' function RANDBETWEEN()
     * 
     * @author Brendan Nolan
     */
    [TestFixture]
    public class TestRandBetween
    {

        private HSSFWorkbook wb;
        private IFormulaEvaluator Evaluator;
        private ICell bottomValueCell;
        private ICell topValueCell;
        private ICell formulaCell;

        [SetUp]
        public void SetUp()
        {
            wb = HSSFTestDataSamples.OpenSampleWorkbook("TestRandBetween.xls");
            Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sheet = wb.CreateSheet("RandBetweenSheet");
            IRow row = sheet.CreateRow(0);
            bottomValueCell = row.CreateCell(0);
            topValueCell = row.CreateCell(1);
            formulaCell = row.CreateCell(2, CellType.Formula);
        }


        protected void tearDown()
        {
            // TODO Auto-generated method stub
        }

        /**
         * Check where values are the same
         */
        [Test]
        public void TestRandBetweenSameValues()
        {

            Evaluator.ClearAllCachedResultValues();
            formulaCell.CellFormula = ("RANDBETWEEN(1,1)");
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(1, formulaCell.NumericCellValue, 0);
            Evaluator.ClearAllCachedResultValues();
            formulaCell.CellFormula = ("RANDBETWEEN(-1,-1)");
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(-1, formulaCell.NumericCellValue, 0);

        }

        /**
         * Check special case where rounded up bottom value is greater than 
         * top value.
         */
        [Test]
        public void TestRandBetweenSpecialCase()
        {


            bottomValueCell.SetCellValue(0.05);
            topValueCell.SetCellValue(0.1);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(1, formulaCell.NumericCellValue, 0);
            bottomValueCell.SetCellValue(-0.1);
            topValueCell.SetCellValue(-0.05);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(0, formulaCell.NumericCellValue, 0);
            bottomValueCell.SetCellValue(-1.1);
            topValueCell.SetCellValue(-1.05);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(-1, formulaCell.NumericCellValue, 0);
            bottomValueCell.SetCellValue(-1.1);
            topValueCell.SetCellValue(-1.1);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(-1, formulaCell.NumericCellValue, 0);
        }

        /**
         * Check top value of BLANK which Excel will Evaluate as 0
         */
        [Test]
        public void TestRandBetweenTopBlank()
        {

            bottomValueCell.SetCellValue(-1);
            topValueCell.SetCellType(CellType.Blank);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.IsTrue(formulaCell.NumericCellValue == 0 || formulaCell.NumericCellValue == -1);

        }
        /**
         * Check where input values are of wrong type
         */
        [Test]
        public void TestRandBetweenWrongInputTypes()
        {
            // Check case where bottom input is of the wrong type
            bottomValueCell.SetCellValue("STRING");
            topValueCell.SetCellValue(1);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(CellType.Error, formulaCell.CachedFormulaResultType);
            Assert.AreEqual(ErrorEval.VALUE_INVALID.ErrorCode, formulaCell.ErrorCellValue);


            // Check case where top input is of the wrong type
            bottomValueCell.SetCellValue(1);
            topValueCell.SetCellValue("STRING");
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(CellType.Error, formulaCell.CachedFormulaResultType);
            Assert.AreEqual(ErrorEval.VALUE_INVALID.ErrorCode, formulaCell.ErrorCellValue);

            // Check case where both inputs are of wrong type
            bottomValueCell.SetCellValue("STRING");
            topValueCell.SetCellValue("STRING");
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(CellType.Error, formulaCell.CachedFormulaResultType);
            Assert.AreEqual(ErrorEval.VALUE_INVALID.ErrorCode, formulaCell.ErrorCellValue);

        }

        /**
         * Check case where bottom is greater than top
         */
        [Test]
        public void TestRandBetweenBottomGreaterThanTop()
        {

            // Check case where bottom is greater than top
            bottomValueCell.SetCellValue(1);
            topValueCell.SetCellValue(0);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(CellType.Error, formulaCell.CachedFormulaResultType);
            Assert.AreEqual(ErrorEval.NUM_ERROR.ErrorCode, formulaCell.ErrorCellValue);
            bottomValueCell.SetCellValue(1);
            topValueCell.SetCellType(CellType.Blank);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(CellType.Error, formulaCell.CachedFormulaResultType);
            Assert.AreEqual(ErrorEval.NUM_ERROR.ErrorCode, formulaCell.ErrorCellValue);
        }

        /**
         * Boundary check of Double MIN and MAX values
         */
        [Test]
        public void TestRandBetweenBoundaryCheck()
        {

            bottomValueCell.SetCellValue(Double.MinValue);
            topValueCell.SetCellValue(Double.MaxValue);
            formulaCell.CellFormula = ("RANDBETWEEN($A$1,$B$1)");
            Evaluator.ClearAllCachedResultValues();
            Evaluator.EvaluateFormulaCell(formulaCell);
            Assert.IsTrue(formulaCell.NumericCellValue >= Double.MinValue && formulaCell.NumericCellValue <= Double.MaxValue);

        }

    }

}
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

namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;

    /**
     * LogicalFunction unit tests.
     */
    [TestFixture]
    public class TestLogicalFunction
    {

        private IFormulaEvaluator Evaluator;
        private IRow row3;
        private ICell cell1;
        private ICell cell2;

        [SetUp]
        public void SetUp()
        {
            IWorkbook wb = new HSSFWorkbook();
            try
            {
                buildWorkbook(wb);
            }
            finally
            {
                wb.Close();
            }
        }

        private void buildWorkbook(IWorkbook wb)
        {
            ISheet sh = wb.CreateSheet();
            IRow row1 = sh.CreateRow(0);
            IRow row2 = sh.CreateRow(1);
            row3 = sh.CreateRow(2);

            row1.CreateCell(0, CellType.Numeric);
            row1.CreateCell(1, CellType.Numeric);

            row2.CreateCell(0, CellType.Numeric);
            row2.CreateCell(1, CellType.Numeric);

            row3.CreateCell(0);
            row3.CreateCell(1);

            CellReference a1 = new CellReference("A1");
            CellReference a2 = new CellReference("A2");
            CellReference b1 = new CellReference("B1");
            CellReference b2 = new CellReference("B2");

            sh.GetRow(a1.Row).GetCell(a1.Col).SetCellValue(35);
            sh.GetRow(a2.Row).GetCell(a2.Col).SetCellValue(0);
            sh.GetRow(b1.Row).GetCell(b1.Col).CellFormula = (/*setter*/"A1/A2");
            sh.GetRow(b2.Row).GetCell(b2.Col).CellFormula = (/*setter*/"NA()");

            Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
        }

        [Test]
        public void TestIsErr()
        {
            cell1 = row3.CreateCell(0);
            cell1.CellFormula = (/*setter*/"ISERR(B1)"); // produces #DIV/0!
            cell2 = row3.CreateCell(1);
            cell2.CellFormula = (/*setter*/"ISERR(B2)"); // produces #N/A

            CellValue cell1Value = Evaluator.Evaluate(cell1);
            CellValue cell2Value = Evaluator.Evaluate(cell2);

            Assert.AreEqual(true, cell1Value.BooleanValue);
            Assert.AreEqual(false, cell2Value.BooleanValue);
        }

        [Test]
        public void TestIsError()
        {
            cell1 = row3.CreateCell(0);
            cell1.CellFormula = (/*setter*/"ISERROR(B1)"); // produces #DIV/0!
            cell2 = row3.CreateCell(1);
            cell2.CellFormula = (/*setter*/"ISERROR(B2)"); // produces #N/A

            CellValue cell1Value = Evaluator.Evaluate(cell1);
            CellValue cell2Value = Evaluator.Evaluate(cell2);

            Assert.AreEqual(true, cell1Value.BooleanValue);
            Assert.AreEqual(true, cell2Value.BooleanValue);
        }
    }

}
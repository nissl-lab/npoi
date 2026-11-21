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
    using NPOI.SS.Formula;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using TestCases.SS.Util;

    /// <summary>
    /// Tests for <see cref="NPOI.SS.Formula.Functions.Or" />
    /// </summary>
    [TestFixture]
    public class TestOrFunction
    {

        private static  readonly OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);

        [Test]
        public void TestMicrosoftExample0()
        {
            //https://support.microsoft.com/en-us/office/or-function-7d17ad14-8700-4281-b308-00b131e22af0
            using(HSSFWorkbook wb = new HSSFWorkbook())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
                HSSFRow row = sheet.CreateRow(0) as HSSFRow;
                HSSFCell cell = row.CreateCell(0) as HSSFCell;
                Utils.AssertBoolean(fe, cell, "OR(TRUE,TRUE)", true);
                Utils.AssertBoolean(fe, cell, "OR(TRUE,FALSE)", true);
                Utils.AssertBoolean(fe, cell, "OR(1=1,2=2,3=3)", true);
                Utils.AssertBoolean(fe, cell, "OR(1=2,2=3,3=4)", false);
            }
        }

        [Test]
        public void TestMicrosoftExample1()
        {
            //https://support.microsoft.com/en-us/office/or-function-7d17ad14-8700-4281-b308-00b131e22af0
            using(HSSFWorkbook wb = InitWorkbook1())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertBoolean(fe, cell, "OR(A2>1,A2<100)", true);
                Utils.AssertDouble(fe, cell, "IF(OR(A2>1,A2<100),A3,\"The value is out of range\")", 100);
                Utils.AssertString(fe, cell, "IF(OR(A2<0,A2>50),A2,\"The value is out of range\")", "The value is out of range");
            }
        }

        [Test]
        public void TestMicrosoftExample2()
        {
            //https://support.microsoft.com/en-us/office/or-function-7d17ad14-8700-4281-b308-00b131e22af0
            using(HSSFWorkbook wb = InitWorkbook2())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(13).CreateCell(3) as HSSFCell;
                Utils.AssertDouble(fe, cell, "IF(OR(B14>=$B$4,C14>=$B$5),B14*$B$6,0)", 314);
            }
        }

        [Test]
        public void TestBug65915()
        {
            //https://bz.apache.org/bugzilla/show_bug.cgi?id=65915
            using(HSSFWorkbook wb = new HSSFWorkbook())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
                HSSFRow row = sheet.CreateRow(0) as HSSFRow;
                HSSFCell cell = row.CreateCell(0) as HSSFCell;
                Utils.AssertDouble(fe, cell, "INDEX({1},1,IF(OR(FALSE,FALSE),1,1))", 1.0);
            }
        }

        [Test]
        public void TestBug65915ArrayFunction()
        {
            //https://bz.apache.org/bugzilla/show_bug.cgi?id=65915
            using(HSSFWorkbook wb = new HSSFWorkbook())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
                HSSFRow row = sheet.CreateRow(0) as HSSFRow;
                HSSFCell cell = row.CreateCell(0) as HSSFCell;
                sheet.SetArrayFormula("INDEX({1},1,IF(OR(FALSE,FALSE),0,1))", new CellRangeAddress(0, 0, 0, 0));
                fe.EvaluateAll();
                ClassicAssert.AreEqual(1.0, cell.NumericCellValue);
            }
        }

        [Test]
        public void TestBug65915Array3Function()
        {
            //https://bz.apache.org/bugzilla/show_bug.cgi?id=65915
            using(HSSFWorkbook wb = new HSSFWorkbook())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
                HSSFRow row = sheet.CreateRow(0) as HSSFRow;
                HSSFCell cell = row.CreateCell(0) as HSSFCell;
                sheet.SetArrayFormula("INDEX({1},1,IF(OR(1=2,2=3,3=4),0,1))", new CellRangeAddress(0, 0, 0, 0));
                fe.EvaluateAll();
                ClassicAssert.AreEqual(1.0, cell.NumericCellValue);
            }
        }

        private static HSSFWorkbook InitWorkbook1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            Utils.AddRow(sheet, 0, "Values");
            Utils.AddRow(sheet, 1, 50);
            Utils.AddRow(sheet, 2, 100);
            return wb;
        }

        private static HSSFWorkbook InitWorkbook2()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            Utils.AddRow(sheet, 0, "Goals");
            Utils.AddRow(sheet, 2, "Criteria", "Amount");
            Utils.AddRow(sheet, 3, "Sales Goal", 8500);
            Utils.AddRow(sheet, 4, "Account Goal", 5);
            Utils.AddRow(sheet, 5, "Commission Rate", 0.02);
            Utils.AddRow(sheet, 6, "Bonus Goal", 12500);
            Utils.AddRow(sheet, 7, "Bonus %", 0.015);
            Utils.AddRow(sheet, 9, "Commission Calculations");
            Utils.AddRow(sheet, 11, "Salesperson", "Total Sales", "Accounts", "Commission", "Bonus");
            Utils.AddRow(sheet, 12, "Millicent Shelton", 10260, 9);
            Utils.AddRow(sheet, 13, "Miguel Ferrari", 15700, 7);
            return wb;
        }
    }
}
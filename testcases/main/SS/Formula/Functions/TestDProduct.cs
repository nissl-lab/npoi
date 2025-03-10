
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
    using NUnit.Framework;
    using TestCases.SS.Util;

    /// <summary>
    /// Testcase for function DPRODUCT()
    /// </summary>
    [TestFixture]
    public class TestDProduct
    {

        //https://support.microsoft.com/en-us/office/dproduct-function-4f96b13e-d49c-47a7-b769-22f6d017cb31
        [Test]
        public void TestMicrosoftExample1()
        {

            using(HSSFWorkbook wb = initWorkbook1(false))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(12) as HSSFCell;
                Utils.AssertDouble(fe, cell, "DPRODUCT(A5:E11, \"Yield\", A1:F3)", 800, 0.0000000001);
                Utils.AssertDouble(fe, cell, "DPRODUCT(A5:E11, \"Yield\", A13:A14)", 720, 0.0000000001);
                Utils.AssertDouble(fe, cell, "DPRODUCT(A5:E11, \"Yield\", B13:C14)", 7560, 0.0000000001);
            }
        }

        [Test]
        public void TestNoMatch()
        {

            using(HSSFWorkbook wb = initWorkbook1(true))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(12) as HSSFCell;
                Utils.AssertDouble(fe, cell, "DPRODUCT(A5:E11, \"Yield\", A1:A2)", 0);
                Utils.AssertDouble(fe, cell, "DPRODUCT(A5:E11, \"Yield\", A1:A3)", 604800);
            }
        }

        private HSSFWorkbook initWorkbook1(bool noMatch)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            Utils.AddRow(sheet, 0, "Tree", "Height", "Age", "Yield", "Profit", "Height");
            if(noMatch)
            {
                Utils.AddRow(sheet, 1, "=NoMatch");
                Utils.AddRow(sheet, 2);
            }
            else
            {
                Utils.AddRow(sheet, 1, "=Apple", ">10", null, null, null, "<16");
                Utils.AddRow(sheet, 2, "=Pear");
            }
            Utils.AddRow(sheet, 3);
            Utils.AddRow(sheet, 4, "Tree", "Height", "Age", "Yield", "Profit");
            Utils.AddRow(sheet, 5, "Apple", 18, 20, 14, 105);
            Utils.AddRow(sheet, 6, "Pear", 12, 12, 10, 96);
            Utils.AddRow(sheet, 7, "Cherry", 13, 14, 9, 105);
            Utils.AddRow(sheet, 8, "Apple", 14, 15, 10, 75);
            Utils.AddRow(sheet, 9, "Pear", 9, 8, 8, 77);
            Utils.AddRow(sheet, 10, "Apple", 8, 9, 6, 45);
            Utils.AddRow(sheet, 11);
            Utils.AddRow(sheet, 12, "Tree", "Height", "Height");
            Utils.AddRow(sheet, 13, "<>Apple", "<>12", "<>9");
            return wb;
        }
    }
}



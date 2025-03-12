
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
    /// Testcase for function DSTDEV() and DSTDEVP()
    /// </summary>
    [TestFixture]
    public class TestDStdev
    {

        //https://support.microsoft.com/en-us/office/dstdev-function-026b8c73-616d-4b5e-b072-241871c4ab96
        [Test]
        public void TestMicrosoftExample1()
        {

            using(HSSFWorkbook wb = initWorkbook1())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(12) as HSSFCell;
                Utils.AssertDouble(fe, cell, "DSTDEV(A5:E11, \"Yield\", A1:A3)", 2.96647939483827, 0.0000000001);
                Utils.AssertDouble(fe, cell, "DSTDEV(A5:E11, \"Yield\", B12:C14)", 2.66458251889485, 0.0000000001);
            }
        }

        //https://support.microsoft.com/en-us/office/dstdevp-function-04b78995-da03-4813-bbd9-d74fd0f5d94b
        [Test]
        public void TestDSTDEVPMicrosoftExample1()
        {

            using(HSSFWorkbook wb = initWorkbook1())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(12) as HSSFCell;
                Utils.AssertDouble(fe, cell, "DSTDEVP(A5:E11, \"Yield\", A1:A3)", 2.65329983228432, 0.0000000001);
                Utils.AssertDouble(fe, cell, "DSTDEVP(A5:E11, \"Yield\", A12:A13)", 0.816496580927726, 0.0000000001);
                Utils.AssertDouble(fe, cell, "DSTDEVP(A5:E11, \"Yield\", B12:C14)", 2.43241991988774, 0.0000000001);
            }
        }

        private HSSFWorkbook initWorkbook1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            Utils.AddRow(sheet, 0, "Tree", "Height", "Age", "Yield", "Profit", "Height");
            Utils.AddRow(sheet, 1, "=Apple", ">10", null, null, null, "<16");
            Utils.AddRow(sheet, 2, "=Pear");
            Utils.AddRow(sheet, 3);
            Utils.AddRow(sheet, 4, "Tree", "Height", "Age", "Yield", "Profit");
            Utils.AddRow(sheet, 5, "Apple", 18, 20, 14, 105);
            Utils.AddRow(sheet, 6, "Pear", 12, 12, 10, 96);
            Utils.AddRow(sheet, 7, "Cherry", 13, 14, 9, 105);
            Utils.AddRow(sheet, 8, "Apple", 14, 15, 10, 75);
            Utils.AddRow(sheet, 9, "Pear", 9, 8, 8, 77);
            Utils.AddRow(sheet, 10, "Apple", 8, 9, 6, 45);
            Utils.AddRow(sheet, 11);
            Utils.AddRow(sheet, 11, "Tree", "Height", "Height");
            Utils.AddRow(sheet, 12, "<>Apple", "<>12", "<>9");
            return wb;
        }
    }
}



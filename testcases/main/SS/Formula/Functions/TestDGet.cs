
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
    using TestCases.SS.Util;


    /// <summary>
    /// Testcase for function DGET()
    /// </summary>
    [TestFixture]
    public class TestDGet
    {

        //https://support.microsoft.com/en-us/office/dget-function-455568bf-4eef-45f7-90f0-ec250d00892e
        [Test]
        public void TestMicrosoftExample1()
        {

            using(HSSFWorkbook wb = initWorkbook1(false, "=Apple"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:A3)", FormulaError.NUM);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", 10);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, 4, A1:F3)", 10);
            }
        }

        [Test]
        public void TestMicrosoftExample1CaseInsensitive()
        {

            using(HSSFWorkbook wb = initWorkbook1(false, "=apple"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:A3)", FormulaError.NUM);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", 10);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, 4, A1:F3)", 10);
            }
        }

        [Test]
        public void TestMicrosoftExample1StartsWith()
        {

            using(HSSFWorkbook wb = initWorkbook1(false, "App"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:A3)", FormulaError.NUM);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", 10);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, 4, A1:F3)", 10);
            }
        }

        [Test]
        public void TestMicrosoftExample1StartsWithLowercase()
        {

            using(HSSFWorkbook wb = initWorkbook1(false, "app"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:A3)", FormulaError.NUM);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", 10);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, 4, A1:F3)", 10);
            }
        }

        [Test]
        public void TestMicrosoftExample1Wildcard()
        {

            using(HSSFWorkbook wb = initWorkbook1(false, "A*le"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:A3)", FormulaError.NUM);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", 10);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, 4, A1:F3)", 10);
            }
        }

        [Test]
        public void TestMicrosoftExample1WildcardLowercase()
        {
            using(HSSFWorkbook wb = initWorkbook1(false, "a*le"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:A3)", FormulaError.NUM);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", 10);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, 4, A1:F3)", 10);
            }
        }

        [Test]
        public void TestMicrosoftExample1AppleWildcardNoMatch()
        {

            using(HSSFWorkbook wb = initWorkbook1(false, "A*x"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:A3)", FormulaError.NUM);
                Utils.AssertError(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", FormulaError.VALUE);
            }
        }

        [Test]
        public void TestMicrosoftExample1Variant()
        {

            using(HSSFWorkbook wb = initWorkbook1(true, "=Apple"))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100) as HSSFCell;
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, \"Yield\", A1:F3)", 6);
                Utils.AssertDouble(fe, cell, "DGET(A5:E11, 4, A1:F3)", 6);
            }
        }

        private HSSFWorkbook initWorkbook1(bool adjustAppleCondition, string appleCondition)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            Utils.AddRow(sheet, 0, "Tree", "Height", "Age", "Yield", "Profit", "Height");
            if(adjustAppleCondition)
            {
                Utils.AddRow(sheet, 1, appleCondition, ">=8", null, null, null, "<12");
            }
            else
            {
                Utils.AddRow(sheet, 1, appleCondition, ">10", null, null, null, "<16");
            }
            Utils.AddRow(sheet, 2, "Pear", ">12");
            Utils.AddRow(sheet, 3);
            Utils.AddRow(sheet, 4, "Tree", "Height", "Age", "Yield", "Profit");
            Utils.AddRow(sheet, 5, "Apple", 18, 20, 14, 105);
            Utils.AddRow(sheet, 6, "Pear", 12, 12, 10, 96);
            Utils.AddRow(sheet, 7, "Cherry", 13, 14, 9, 105);
            Utils.AddRow(sheet, 8, "Apple", 14, null, 10, 75);
            Utils.AddRow(sheet, 9, "Pear", 9, 8, 8, 77);
            Utils.AddRow(sheet, 10, "Apple", 8, 9, 6, 45);
            return wb;
        }
    }
}


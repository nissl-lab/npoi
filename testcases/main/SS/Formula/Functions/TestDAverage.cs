
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using TestCases.SS.Util;

    /// <summary>
    /// Testcase for function DAVERAGE()
    /// </summary>
    [TestFixture]
    public class TestDAverage
    {

        //https://support.microsoft.com/en-us/office/daverage-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee
        [Test]
        public void TestMicrosoftExample1()
        {

            using(HSSFWorkbook wb = initWorkbook1())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                HSSFCell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(12) as HSSFCell;
                Utils.AssertDouble(fe, cell, "DAVERAGE(A4:E10, \"Yield\", A1:B2)", 12);
                Utils.AssertDouble(fe, cell, "DAVERAGE(A4:E10, 3, A4:E10)", 13);
                Utils.AssertDouble(fe, cell, "DAVERAGE(A4:E10, \"Profit\", A12:A13)", 92.6);
                Utils.AssertDouble(fe, cell, "DAVERAGE(A4:E10, \"Profit\", B12:C13)", 82.5);
            }
        }

        private HSSFWorkbook initWorkbook1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            Utils.AddRow(sheet, 0, "Tree", "Height", "Age", "Yield", "Profit", "Height");
            Utils.AddRow(sheet, 1, "=Apple", ">10", null, null, null, "<16");
            Utils.AddRow(sheet, 2, "=Pear");
            Utils.AddRow(sheet, 3, "Tree", "Height", "Age", "Yield", "Profit");
            Utils.AddRow(sheet, 4, "Apple", 18, 20, 14, 105);
            Utils.AddRow(sheet, 5, "Pear", 12, 12, 10, 96);
            Utils.AddRow(sheet, 6, "Cherry", 13, 14, 9, 105);
            Utils.AddRow(sheet, 7, "Apple", 14, 15, 10, 75);
            Utils.AddRow(sheet, 8, "Pear", 9, 8, 8, 76.8);
            Utils.AddRow(sheet, 9, "Apple", 8, 9, 6, 45);
            Utils.AddRow(sheet, 10);
            Utils.AddRow(sheet, 11, "Tree", "Height", "Height");
            Utils.AddRow(sheet, 12, "<>Apple", "<>12", "<>9");
            return wb;
        }
    }
}



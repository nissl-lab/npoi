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

namespace TestCases.SS.Formula.PTG
{

    using System;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;
    /**
     * Tests for functions from external workbooks (e.g. YEARFRAC).
     * 
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestExternalFunctionFormulas
    {

        /**
         * Tests <c>NameXPtg.ToFormulaString(Workbook)</c> and logic in Workbook below that   
         */
        [Test]
        public void TestReadFormulaContainingExternalFunction()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("externalFunctionExample.xls");

            String expectedFormula = "YEARFRAC(B1,C1)";
            ISheet sht = wb.GetSheetAt(0);
            String cellFormula = sht.GetRow(0).GetCell(0).CellFormula;
            Assert.AreEqual(expectedFormula, cellFormula);
        }
        [Test]
        public void TestParse()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("externalFunctionExample.xls");
            Ptg[] ptgs = HSSFFormulaParser.Parse("YEARFRAC(B1,C1)", wb);
            Assert.AreEqual(4, ptgs.Length);
            Assert.AreEqual(typeof(NameXPtg), ptgs[0].GetType());

            wb.GetSheetAt(0).GetRow(0).CreateCell(6).CellFormula = ("YEARFRAC(C1,B1)");
#if !HIDE_UNREACHABLE_CODE
            if (false)
            {
                // In case you fancy Checking in excel
                try
                {
                    FileStream tempFile = File.Create("testExtFunc.xls");
                    //FileOutputStream fout = new FileOutputStream(tempFile);
                    wb.Write(tempFile);
                    tempFile.Close();
                    //Console.WriteLine("check out " + tempFile.getAbsolutePath());
                }
                catch (IOException e)
                {
                    throw e;
                }
            }
#endif
        }
        [Test]
        public void TestEvaluate()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US"); 
            
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("externalFunctionExample.xls");
            ISheet sheet = wb.GetSheetAt(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ConfirmCellEval(sheet, 0, 0, fe, "YEARFRAC(B1,C1)", 29.0 / 90.0);
            ConfirmCellEval(sheet, 1, 0, fe, "YEARFRAC(B2,C2)", 0.0);
            ConfirmCellEval(sheet, 2, 0, fe, "YEARFRAC(B3,C3,D3)", 0.0);
            ConfirmCellEval(sheet, 3, 0, fe, "IF(ISEVEN(3),1.2,1.6)", 1.6);
            ConfirmCellEval(sheet, 4, 0, fe, "IF(ISODD(3),1.2,1.6)", 1.2);
        }

        private static void ConfirmCellEval(ISheet sheet, int rowIx, int colIx,
                HSSFFormulaEvaluator fe, String expectedFormula, double expectedResult)
        {
            ICell cell = sheet.GetRow(rowIx).GetCell(colIx);
            Assert.AreEqual(expectedFormula, cell.CellFormula);
            CellValue cv = fe.Evaluate(cell);
            Assert.AreEqual(expectedResult, cv.NumberValue, 0.0);
        }
    }

}
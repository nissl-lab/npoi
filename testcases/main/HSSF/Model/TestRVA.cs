/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Model
{
    using System;
    using System.Text;
    using NUnit.Framework;

    using TestCases.HSSF;
    using NPOI.SS.Formula;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using TestCases.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Formula.PTG;

    /**
     * Tests 'operand class' transformation performed by
     * <c>OperandClassTransformer</c> by comparing its results with those
     * directly produced by Excel (in a sample spReadsheet).
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRVA
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void TestFormulasRVA()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("testRVA.xls");
            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);

            int countFailures = 0;
            int countErrors = 0;

            int rowIx = 0;
            while (rowIx < 65535)
            {
                IRow row = sheet.GetRow(rowIx);
                if (row == null)
                {
                    break;
                }
                ICell cell = row.GetCell(0);
                if (cell == null || cell.CellType == NPOI.SS.UserModel.CellType.Blank)
                {
                    break;
                }
                String formula = cell.CellFormula;
                try
                {
                    ConfirmCell(cell, formula, wb);
                }
                catch (AssertionException e)
                {
                    Console.Error.WriteLine("Problem with row[" + rowIx + "] formula '" + formula + "'");
                    Console.Error.WriteLine(e.Message);
                    countFailures++;
                }
                catch (Exception)
                {
                    Console.Error.WriteLine("Problem with row[" + rowIx + "] formula '" + formula + "'");
                    countErrors++;
                }
                rowIx++;
            }
            if (countErrors + countFailures > 0)
            {
                String msg = "One or more RVA tests failed: countFailures=" + countFailures
                        + " countFailures=" + countErrors + ". See stderr for details.";
                throw new AssertionException(msg);
            }
        }

        private void ConfirmCell(ICell formulaCell, String formula, HSSFWorkbook wb)
        {
            Ptg[] excelPtgs = FormulaExtractor.GetPtgs(formulaCell);
            Ptg[] poiPtgs = HSSFFormulaParser.Parse(formula, wb);
            int nExcelTokens = excelPtgs.Length;
            int nPoiTokens = poiPtgs.Length;
            if (nExcelTokens != nPoiTokens)
            {
                if (nExcelTokens == nPoiTokens + 1 && excelPtgs[0].GetType() == typeof(AttrPtg))
                {
                    // compensate for missing tAttrVolatile, which belongs in any formula 
                    // involving OFFSET() et al. POI currently does not insert where required
                    Ptg[] temp = new Ptg[nExcelTokens];
                    temp[0] = excelPtgs[0];
                    Array.Copy(poiPtgs, 0, temp, 1, nPoiTokens);
                    poiPtgs = temp;
                }
                else
                {
                    throw new Exception("Expected " + nExcelTokens + " tokens but got "
                            + nPoiTokens);
                }
            }
            bool hasMismatch = false;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < nExcelTokens; i++)
            {
                Ptg poiPtg = poiPtgs[i];
                Ptg excelPtg = excelPtgs[i];
                if (excelPtg.GetType() != poiPtg.GetType())
                {
                    hasMismatch = true;
                    sb.Append("  mismatch token type[" + i + "] " + GetShortClassName(excelPtg) + " "
                            + excelPtg.RVAType + " - " + GetShortClassName(poiPtg) + " "
                            + poiPtg.RVAType);
                    sb.Append(Environment.NewLine);
                    continue;
                }
                if (poiPtg.IsBaseToken)
                {
                    continue;
                }
                sb.Append("  token[" + i + "] " + excelPtg.ToString() + " "
                        + excelPtg.RVAType);

                if (excelPtg.PtgClass != poiPtg.PtgClass)
                {
                    hasMismatch = true;
                    sb.Append(" - was " + poiPtg.RVAType);
                }
                sb.Append(Environment.NewLine);
            }
            //if (false)
            //{ // Set 'true' to see trace of RVA values
            //    Console.WriteLine(formula);
            //    Console.WriteLine(sb.ToString());
            //}
            if (hasMismatch)
            {
                throw new AssertionException(sb.ToString());
            }
        }

        private String GetShortClassName(Object o)
        {
            String cn = o.GetType().Name;
            int pos = cn.LastIndexOf('.');
            return cn.Substring(pos + 1);
        }
       
    }
}
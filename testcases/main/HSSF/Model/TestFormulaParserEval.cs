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
    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;

    /**
     * Test the low level formula parser functionality,
     *  but using parts which need to use 
     *  HSSFFormulaEvaluator.
     */
    [TestFixture]
    public class TestFormulaParserEval
    {
        [Test]
        public void TestWithNamedRange()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();

            ISheet s = workbook.CreateSheet("Foo");
            s.CreateRow(0).CreateCell((short)0).SetCellValue(1.1);
            s.CreateRow(1).CreateCell((short)0).SetCellValue(2.3);
            s.CreateRow(2).CreateCell((short)2).SetCellValue(3.1);

            IName name = workbook.CreateName();
            name.NameName = ("testName");
            name.RefersToFormula = ("A1:A2");

            ConfirmParseFormula(workbook);

            // Now make it a single cell
            name.RefersToFormula = ("C3");
            ConfirmParseFormula(workbook);

            // And make it non-contiguous
            name.RefersToFormula = ("A1:A2,C3");
            ConfirmParseFormula(workbook);
        }
        	/**
	 * Makes sure that a formula referring to the named range parses properly
	 */
        private static void ConfirmParseFormula(HSSFWorkbook workbook)
        {
            Ptg[] ptgs = HSSFFormulaParser.Parse("SUM(testName)", workbook);
            Assert.IsTrue(ptgs.Length == 2, "two tokens expected, got " + ptgs.Length);
            Assert.AreEqual(typeof(NamePtg), ptgs[0].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[1].GetType());
        }

        [Test]
        public void TestEvaluateFormulaWithRowBeyond32768_Bug44539()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            cell.CellFormula = ("SUM(A32769:A32770)");

            // put some values in the cells to make the evaluation more interesting
            sheet.CreateRow(32768).CreateCell((short)0).SetCellValue(31);
            sheet.CreateRow(32769).CreateCell((short)0).SetCellValue(11);

            //HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(sheet, wb);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            NPOI.SS.UserModel.CellValue result;
            try
            {
                result = fe.Evaluate(cell);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Found reference to named range \"A\", but that named range wasn't defined!"))
                {
                    Assert.Fail("Identifed bug 44539");
                }
                throw;
            }
            Assert.AreEqual(NPOI.SS.UserModel.CellType.Numeric, result.CellType);
            Assert.AreEqual(42.0, result.NumberValue, 0.0);
        }
    }
}
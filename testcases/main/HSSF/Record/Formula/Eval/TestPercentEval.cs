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

namespace TestCases.SS.Formula.Eval
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.SS.Formula;
    using TestCases.SS.Formula.Functions;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;

    /**
     * Test for percent operator evaluator.
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestPercentEval
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [ClassInitialize()]
        public static void PrepareCultere(TestContext testContext) 
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        private void Confirm(ValueEval arg, double expectedResult)
        {
            ValueEval[] args = { 
			arg,	
		};

            NPOI.SS.Formula.Functions.Function opEval = PercentEval.instance;
            double result = NumericFunctionInvoker.Invoke(opEval, args, -1, (short)-1);

            Assert.AreEqual(expectedResult, result, 0);
        }
        [TestMethod]
        public void TestBasic()
        {
            Confirm(new NumberEval(5), 0.05);
            Confirm(new NumberEval(3000), 30.0);
            Confirm(new NumberEval(-150), -1.5);
            Confirm(new StringEval("0.2"), 0.002);
            Confirm(BoolEval.TRUE, 0.01);
        }
        [TestMethod]
        public void TestInSpreadSheet()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            cell.CellFormula = ("B1%");
            row.CreateCell(1).SetCellValue(50.0);

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(sheet, wb);
            NPOI.SS.UserModel.CellValue cv;
            try
            {
                cv = fe.Evaluate(cell);
            }
            catch (Exception e)
            {
                if (e.InnerException is NullReferenceException)
                {
                    throw new AssertFailedException("Identified bug 44608");
                }
                // else some other unexpected error
                throw e;
            }
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(0.5, cv.NumberValue, 0.0);
        }

    }
}

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

namespace TestCases.SS.Formula.Eval
{

    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using TestCases.SS.Formula.Functions;

    /**
     * Test for percent operator Evaluator.
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestPercentEval
    {

        private static void Confirm(ValueEval arg, double expectedResult)
        {
            ValueEval[] args = {
			arg,
		};

            double result = NumericFunctionInvoker.Invoke(PercentEval.instance, args, 0, 0);

            Assert.AreEqual(expectedResult, result, 0);
        }
        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            Confirm(new NumberEval(5), 0.05);
            Confirm(new NumberEval(3000), 30.0);
            Confirm(new NumberEval(-150), -1.5);
            Confirm(new StringEval("0.2"), 0.002);
            Confirm(BoolEval.TRUE, 0.01);
        }
        [Test]
        public void Test1x1Area()
        {
            AreaEval ae = EvalFactory.CreateAreaEval("B2:B2", new ValueEval[] { new NumberEval(50), });
            Confirm(ae, 0.5);
        }
        [Test]
        public void TestInSpreadSheet()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.CellFormula=("B1%");
            row.CreateCell(1).SetCellValue(50.0);

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv;
            try
            {
                cv = fe.Evaluate(cell);
            }
            catch (SystemException e)
            {
                if (e.InnerException is NullReferenceException)
                {
                    throw new AssertionException("Identified bug 44608");
                }
                // else some other unexpected error
                throw e;
            }
            Assert.AreEqual(CellType.Numeric, cv.CellType);
            Assert.AreEqual(0.5, cv.NumberValue, 0.0);
        }
    }

}
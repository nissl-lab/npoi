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


namespace TestCases.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Tests for {@link FinanceFunction#NPER}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestNper
    {
        [Test]
        public void TestSimpleEvaluate()
        {

            ValueEval[] args = {
			new NumberEval(0.05),
			new NumberEval(250),
			new NumberEval(-1000),
		};
            ValueEval result = FinanceFunction.NPER.Evaluate(args, 0, (short)0);

            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(4.57353557, ((NumberEval)result).NumberValue, 0.00000001);
        }
        [Test]
        public void TestEvaluate_bug_45732()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            ICell cell = sheet.CreateRow(0).CreateCell(0);

            cell.CellFormula = ("NPER(12,4500,100000,100000)");
            cell.SetCellValue(15.0);
            Assert.AreEqual("NPER(12,4500,100000,100000)", cell.CellFormula);
            Assert.AreEqual(CellType.Numeric, cell.CachedFormulaResultType);
            Assert.AreEqual(15.0, cell.NumericCellValue, 0.0);

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            fe.EvaluateFormulaCell(cell);
            Assert.AreEqual(CellType.Error, cell.CachedFormulaResultType);
            Assert.AreEqual(FormulaError.NUM.Code, cell.ErrorCellValue);
            wb.Close();
        }
    }

}
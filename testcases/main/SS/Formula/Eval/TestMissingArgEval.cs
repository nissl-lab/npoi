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
    using NPOI.SS.UserModel;

    /**
     * Tests for {@link MissingArgEval}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestMissingArgEval
    {
        [Test]
        public void TestEvaluateMissingArgs()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ISheet sheet = wb.CreateSheet("Sheet1");
            ICell cell = sheet.CreateRow(0).CreateCell(0);

            cell.CellFormula=("if(true,)");
            fe.ClearAllCachedResultValues();
            CellValue cv;
            try
            {
                cv = fe.Evaluate(cell);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                throw new AssertionException("Missing args Evaluation not implemented (bug 43354");
            }
            // MissingArg -> BlankEval -> zero (as formula result)
            Assert.AreEqual(0.0, cv.NumberValue, 0.0);

            // MissingArg -> BlankEval -> empty string (in concatenation)
            cell.CellFormula=("\"abc\"&if(true,)");
            fe.ClearAllCachedResultValues();
            Assert.AreEqual("abc", fe.Evaluate(cell).StringValue);
        }
        [Test]
        public void TestCompareMissingArgs()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ISheet sheet = wb.CreateSheet("Sheet1");
            ICell cell = sheet.CreateRow(0).CreateCell(0);

            cell.SetCellFormula("iferror(0/0,)<0");
            fe.ClearAllCachedResultValues();
            CellValue cv = fe.Evaluate(cell);
            Assert.IsFalse(cv.BooleanValue);

            cell.SetCellFormula("iferror(0/0,)<=0");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsTrue(cv.BooleanValue);

            cell.SetCellFormula("iferror(0/0,)=0");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsTrue(cv.BooleanValue);

            cell.SetCellFormula("iferror(0/0,)>=0");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsTrue(cv.BooleanValue);

            cell.SetCellFormula("iferror(0/0,)>0");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsFalse(cv.BooleanValue);

            // invert above for code coverage
            cell.SetCellFormula("0<iferror(0/0,)");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsFalse(cv.BooleanValue);

            cell.SetCellFormula("0<=iferror(0/0,)");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsTrue(cv.BooleanValue);

            cell.SetCellFormula("0=iferror(0/0,)");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsTrue(cv.BooleanValue);

            cell.SetCellFormula("0>=iferror(0/0,)");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsTrue(cv.BooleanValue);

            cell.SetCellFormula("0>iferror(0/0,)");
            fe.ClearAllCachedResultValues();
            cv = fe.Evaluate(cell);
            Assert.IsFalse(cv.BooleanValue);
        }
        [Test]
        public void TestCountFuncs()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ISheet sheet = wb.CreateSheet("Sheet1");
            ICell cell = sheet.CreateRow(0).CreateCell(0);

            cell.CellFormula=("COUNT(C5,,,,)"); // 4 missing args, C5 is blank 
            Assert.AreEqual(4.0, fe.Evaluate(cell).NumberValue, 0.0);

            cell.CellFormula=("COUNTA(C5,,)"); // 2 missing args, C5 is blank 
            fe.ClearAllCachedResultValues();
            Assert.AreEqual(2.0, fe.Evaluate(cell).NumberValue, 0.0);
        }
    }
}

/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NUnit.Framework;

namespace TestCases.SS.Formula
{
    /**
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestFunctionRegistry
    {
        [Test]
        public void TestRegisterInRuntime()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet("Sheet1");
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            HSSFCell cellA = (HSSFCell)row.CreateCell(0);
            cellA.CellFormula = ("FISHER(A5)");
            CellValue cv;
            try
            {
                //NPOI
                //Run it twice in NUnit Gui Window, the first passed but the second failed.
                //Maybe the function was cached. Ignore it.
                cv = fe.Evaluate(cellA);
                Assert.Fail("expectecd exception");
            }
            catch (NotImplementedException)
            {
                ;
            }

            FunctionEval.RegisterFunction("FISHER", new Function1());/*Function() {
            public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
                return ErrorEval.NA;
            }
        });*/

            cv = fe.Evaluate(cellA);
            Assert.AreEqual(ErrorEval.NA.ErrorCode, cv.ErrorValue);

            HSSFCell cellB = (HSSFCell)row.CreateCell(1);
            cellB.CellFormula = ("CUBEMEMBERPROPERTY(A5)");
            try
            {
                cv = fe.Evaluate(cellB);
                Assert.Fail("expectecd exception");
            }
            catch (NotImplementedException)
            {
                ;
            }

            AnalysisToolPak.RegisterFunction("CUBEMEMBERPROPERTY", new FreeRefFunction1());/*FreeRefFunction() {
            public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
                return ErrorEval.NUM_ERROR;
            }
        });*/

            cv = fe.Evaluate(cellB);
            Assert.AreEqual(ErrorEval.NUM_ERROR.ErrorCode, cv.ErrorValue);
        }

        private class Function1 : NPOI.SS.Formula.Functions.Function
        {
            public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
            {
                return ErrorEval.NA;
            }
        }

        private class FreeRefFunction1 : FreeRefFunction
        {
            public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
            {
                return ErrorEval.NUM_ERROR;
            }
        }

        class Function2 : NPOI.SS.Formula.Functions.Function
        {
            public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
            {
                return ErrorEval.NA;
            }
        }
        [Test]
        public void TestExceptions()
        {
            NPOI.SS.Formula.Functions.Function func = new Function2();
            try
            {
                FunctionEval.RegisterFunction("SUM", func);
                Assert.Fail("expectecd exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("POI already implememts SUM" +
                        ". You cannot override POI's implementations of Excel functions", e.Message);
            }
            try
            {
                FunctionEval.RegisterFunction("SUMXXX", func);
                Assert.Fail("expectecd exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Unknown function: SUMXXX", e.Message);
            }
            try
            {
                FunctionEval.RegisterFunction("ISODD", func);
                Assert.Fail("expectecd exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("ISODD is a function from the Excel Analysis Toolpack. " +
                        "Use AnalysisToolpack.RegisterFunction(String name, FreeRefFunction func) instead.", e.Message);
            }

            FreeRefFunction atpFunc = new FreeRefFunction2();/*FreeRefFunction() {
            public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
                return ErrorEval.NUM_ERROR;
            }
        };*/
            try
            {
                AnalysisToolPak.RegisterFunction("ISODD", atpFunc);
                Assert.Fail("expectecd exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("POI already implememts ISODD" +
                        ". You cannot override POI's implementations of Excel functions", e.Message);
            }
            try
            {
                AnalysisToolPak.RegisterFunction("ISODDXXX", atpFunc);
                Assert.Fail("expectecd exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("ISODDXXX is not a function from the Excel Analysis Toolpack.", e.Message);
            }
            try
            {
                AnalysisToolPak.RegisterFunction("SUM", atpFunc);
                Assert.Fail("expectecd exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("SUM is a built-in Excel function. " +
                        "Use FunctoinEval.RegisterFunction(String name, Function func) instead.", e.Message);
            }
        }
        class FreeRefFunction2 : FreeRefFunction
        {
            public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
            {
                return ErrorEval.NUM_ERROR;
            }
        }
    }

}
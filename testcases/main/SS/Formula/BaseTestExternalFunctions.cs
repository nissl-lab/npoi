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
namespace TestCases.SS.Formula
{

    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.UDF;
    using NPOI.SS.UserModel;
    using TestCases.SS;

    /**
     * Test Setting / Evaluating of Analysis Toolpack and user-defined functions
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class BaseTestExternalFunctions
    {
        protected ITestDataProvider _testDataProvider;
        private String atpFile;
        public BaseTestExternalFunctions()
        {
            _testDataProvider = TestCases.HSSF.HSSFITestDataProvider.Instance;
            
        }

        /**
        * @param TestDataProvider an object that provides Test data in HSSF / XSSF specific way
        */
        protected BaseTestExternalFunctions(ITestDataProvider TestDataProvider, String atpFile)
        {
            _testDataProvider = TestDataProvider;
            this.atpFile = atpFile;
        }
        [Test]
        public void TestExternalFunctions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();

            ICell cell1 = sh.CreateRow(0).CreateCell(0);
            // functions from the Excel Analysis Toolpack
            cell1.CellFormula=("ISODD(1)+ISEVEN(2)"); 
            Assert.AreEqual("ISODD(1)+ISEVEN(2)", cell1.CellFormula);

            ICell cell2 = sh.CreateRow(1).CreateCell(0);
            //unregistered functions are parseable and renderable, but may not be evaluateable
            cell2.SetCellFormula("MYFUNC(\"B1\")"); 

            try
            {
                evaluator.Evaluate(cell2);
                Assert.Fail("Expected NotImplementedFunctionException/NotImplementedException");
            }
            catch (NotImplementedException e)
            {
                if (!(e.InnerException is NotImplementedFunctionException))
                    throw e;
                // expected
                // Alternatively, a future implementation of evaluate could return #NAME? error to align behavior with Excel
                // assertEquals(ErrorEval.NAME_INVALID, ErrorEval.valueOf(evaluator.evaluate(cell2).getErrorValue()));
            }
            //try
            //{
            //    //NPOI
            //    //Run it twice in NUnit Gui Window, the first passed but the second failed.
            //    //Maybe the function was cached. Ignore it.
            //    cell2.CellFormula=("MYBASEEXTFUNC(\"B1\")");
            //    Assert.Fail("Should fail because MYBASEEXTFUNC is an unknown function");
            //}
            //catch (FormulaParseException)
            //{
            //    ; //expected
            //}

            wb.AddToolPack(customToolpack);

            cell2.CellFormula = ("MYBASEEXTFUNC(\"B1\")");
            Assert.AreEqual("MYBASEEXTFUNC(\"B1\")", cell2.CellFormula);

            ICell cell3 = sh.CreateRow(2).CreateCell(0);
            cell3.CellFormula = ("MYBASEEXTFUNC2(\"C1\")&\"-\"&A2");  //where A2 is defined above
            Assert.AreEqual("MYBASEEXTFUNC2(\"C1\")&\"-\"&A2", cell3.CellFormula);

            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual(2.0, Evaluator.Evaluate(cell1).NumberValue);
            Assert.AreEqual("B1abc", Evaluator.Evaluate(cell2).StringValue);
            Assert.AreEqual("C1abc2-B1abc", Evaluator.Evaluate(cell3).StringValue);

            wb.Close();
        }

        /**
         * Test invoking saved ATP functions
         *
         * @param TestFile  either atp.xls or atp.xlsx
         */
        public void BaseTestInvokeATP(String testFile)
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(testFile);
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.GetSheetAt(0);
            // these two are not imlemented in r
            Assert.AreEqual("DELTA(1.3,1.5)", sh.GetRow(0).GetCell(1).CellFormula);
            Assert.AreEqual("COMPLEX(2,4)", sh.GetRow(1).GetCell(1).CellFormula);

            ICell cell2 = sh.GetRow(2).GetCell(1);
            Assert.AreEqual("ISODD(2)", cell2.CellFormula);
            Assert.AreEqual(false, Evaluator.Evaluate(cell2).BooleanValue);
            Assert.AreEqual(CellType.Boolean, Evaluator.EvaluateFormulaCell(cell2));

            ICell cell3 = sh.GetRow(3).GetCell(1);
            Assert.AreEqual("ISEVEN(2)", cell3.CellFormula);
            Assert.AreEqual(true, Evaluator.Evaluate(cell3).BooleanValue);
            Assert.AreEqual(CellType.Boolean, Evaluator.EvaluateFormulaCell(cell3));

            wb.Close();
        }

        // define two custom user-defined functions
        private class MyBaseExtFunc : FreeRefFunction
        {
            public MyBaseExtFunc()
            {
                //
            }

            public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
            {
                if (args.Length != 1 || !(args[0] is StringEval))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                StringEval input = (StringEval)args[0];
                return new StringEval(input.StringValue + "abc");
            }
        }

        private class MyBaseExtFunc2 : FreeRefFunction
        {
            public MyBaseExtFunc2()
            {
                //
            }

            public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
            {
                if (args.Length != 1 || !(args[0] is StringEval))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                StringEval input = (StringEval)args[0];
                return new StringEval(input.StringValue + "abc2");
            }
        }

        /**
         * register the two Test UDFs in a UDF Finder, to be passed to the workbook MYBASEEXTFUNC
         */
        private static UDFFinder customToolpack = new DefaultUDFFinder(
                new String[] { "myBaseExtFunc", "myBaseExtFunc2" },
                new FreeRefFunction[] { new MyBaseExtFunc(), new MyBaseExtFunc2() }
        );

    }

}
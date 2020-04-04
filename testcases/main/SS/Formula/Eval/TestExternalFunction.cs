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
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.UDF;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;

    /**
     * @author Josh Micich
     * @author Petr Udalau - registering UDFs in workbook and using ToolPacks.
     */
    [TestFixture]
    public class TestExternalFunction
    {

        private class MyFunc : FreeRefFunction
        {
            public MyFunc()
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

        private class MyFunc2 : FreeRefFunction
        {
            public MyFunc2()
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
         * Checks that an external function can Get invoked from the formula
         * Evaluator.
         */
        [Test]
        public void TestInvoke()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("testNames.xls");
            ISheet sheet = wb.GetSheetAt(0);

            /**
             * register the two Test UDFs in a UDF Finder, to be passed to the Evaluator
             */
            UDFFinder udff1 = new DefaultUDFFinder(new String[] { "myFunc", },
                    new FreeRefFunction[] { new MyFunc(), });
            UDFFinder udff2 = new DefaultUDFFinder(new String[] { "myFunc2", },
                    new FreeRefFunction[] { new MyFunc2(), });
            UDFFinder udff = new AggregatingUDFFinder(udff1, udff2);


            IRow row = sheet.GetRow(0);
            ICell myFuncCell = row.GetCell(1); // =myFunc("_")

            ICell myFunc2Cell = row.GetCell(2); // =myFunc2("_")

            HSSFFormulaEvaluator fe = HSSFFormulaEvaluator.Create(wb, null, udff);
            Assert.AreEqual("_abc", fe.Evaluate(myFuncCell).StringValue);
            Assert.AreEqual("_abc2", fe.Evaluate(myFunc2Cell).StringValue);
        }
    }

}
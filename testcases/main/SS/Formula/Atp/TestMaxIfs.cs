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

namespace TestCases.SS.Formula.Atp
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Atp;
    using NUnit.Framework;
    using TestCases.SS.Formula.Functions;

    /**
     * Test cases for MAXIFS()
     */
    [TestFixture]
    public class TestMaxIfs
    {
        private static OperationEvaluationContext EC = new OperationEvaluationContext(null, null, 0, 1, 0, null);

        private static ValueEval InvokeMaxIfs(ValueEval[] args, OperationEvaluationContext ec)
        {
            return new Maxifs().Evaluate(args, EC);
        }

        private static void ConfirmDouble(double expected, ValueEval actualEval)
        {
            if (!(actualEval is NumericValueEval))
            {
                throw new AssertionException("Expected numeric result");
            }
            NumericValueEval nve = (NumericValueEval)actualEval;
            Assert.AreEqual(expected, nve.NumberValue, 0);
        }

        private static void Confirm(double expectedResult, ValueEval[] args)
        {
            ConfirmDouble(expectedResult, InvokeMaxIfs(args, EC));
        }

        /**
         *  Example 1 from
         *  https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883
         */
        [Test]
        public void TestExample1()
        {
            // mimic test sample from https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883
            ValueEval[] a2a7 = new ValueEval[]
            {
                new NumberEval(89),
                new NumberEval(93),
                new NumberEval(91),
                new NumberEval(96),
                new NumberEval(85),
                new NumberEval(88)
            };

            ValueEval[] b2b7 = new ValueEval[]
            {
                new NumberEval(1),
                new NumberEval(2),
                new NumberEval(2),
                new NumberEval(3),
                new NumberEval(1),
                new NumberEval(1)
            };

            // "=MAXIFS(A2:A7,B2:B7,1)"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A7", a2a7),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new NumberEval(1)
            };
            Confirm(89.0, args);

            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A7", a2a7),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new StringEval(">1")
            };
            Confirm(96.0, args);

            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A7", a2a7),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new StringEval(">1"),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new StringEval("<3")
            };
            Confirm(93.0, args);
        }

         /**
         *  Ensure that this works with non-numeric data within the processed values.
         */
        [Test]
        public void TestMinWithNonNumeric()
        {
            ValueEval[] a2a7 = new ValueEval[]
            {
                new NumberEval(-89),
                new NumberEval(-93),
                new NumberEval(-96),
                new NumberEval(-85),
                new StringEval("Test"),
                new NumberEval(-88)
            };

            ValueEval[] b2b7 = new ValueEval[]
            {
                new NumberEval(1),
                new NumberEval(2),
                new NumberEval(2),
                new NumberEval(3),
                new NumberEval(1),
                new NumberEval(1)
            };

            // "=MaxIFS(A2:A7, B2:B7, "1")"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A7", a2a7),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new NumberEval(1)
            };
            Confirm(-88.00, args);

            // "=MaxIFS(A2:A7, B2:B7, ">1")"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A7", a2a7),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new StringEval(">1")
            };
            Confirm(-85.0, args);

            // "=MaxIFS(A2:A7, B2:B7, ">1", B2:B7, "<3")"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A7", a2a7),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new StringEval(">1"),
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                new StringEval("<3")
            };
            Confirm(-93.0, args);
        }
    }
}
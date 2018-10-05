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

using NPOI.SS.Formula;

namespace TestCases.SS.Formula.Functions {
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;

    using NUnit.Framework;

    /**
     * Test cases for SUM
     *
     * PRODUCT()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestAverageif
    {
        private static NumberEval _30 = new NumberEval(30);
        private static NumberEval _40 = new NumberEval(40);
        private static NumberEval _50 = new NumberEval(50);
        private static NumberEval _60 = new NumberEval(60);

        private static OperationEvaluationContext EC = new OperationEvaluationContext(null, null, 0, 1, 0, null);

        private static ValueEval InvokeAverageif(ValueEval[] args, OperationEvaluationContext ec)
        {
            return new AverageIf().Evaluate(args, ec);
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

        private static void Confirm(double expectedResult, ValueEval[] args) {
            ConfirmDouble(expectedResult, InvokeAverageif(args, EC));
        }

        /**
         *  Example 1 from
         *  https://support.office.com/en-us/article/averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642
         */
        [Test]
        public void TestExample1() {
            ValueEval[] b2b5 = new ValueEval[]
            {
                new NumberEval(7000),
                new NumberEval(14000),
                new NumberEval(21000),
                new NumberEval(28000),
            };

            ValueEval[] a2a5 = new ValueEval[]
            {
                new NumberEval(100000),
                new NumberEval(200000),
                new NumberEval(300000),
                new NumberEval(400000),
            };

            // "=AVERAGEIF(B2:B5,"<23000")" 
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("B2:B5", b2b5),
                new StringEval("<23000"),
            };
            Confirm(14000, args);

            // "=AVERAGEIF(A2:A5,"<250000")"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A5", a2a5),
                new StringEval("<250000"),
            };
            Confirm(150000, args);

            //"=AVERAGEIF(A2:A5,"<95000")"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A5", a2a5),
                new StringEval("<95000"),
            };
            Assert.AreEqual(ErrorEval.VALUE_INVALID, InvokeAverageif(args, EC));

            // "=AVERAGEIF(A2:A5,">250000",B2:B5)"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A5", a2a5),
                new StringEval(">250000"),
                EvalFactory.CreateAreaEval("B2:B5", b2b5),
            };
                Confirm(24500, args);
        }

        /**
         *  Example 2 from
         *  https://support.office.com/en-us/article/averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642
         */
        [Test]
        public void TestExample2() {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            ValueEval[] b2b6 = new ValueEval[]
            {
                new NumberEval(45678),
                new NumberEval(23789),
                new NumberEval(-4789),
                new NumberEval(0),
                new NumberEval(9678),
            };

            ValueEval[] a2a6 = new ValueEval[]
            {
                new StringEval("East"),
                new StringEval("West"),
                new StringEval("North"),
                new StringEval("South(New Office)"),
                new StringEval("MidWest"),
            };

            // "=AVERAGEIF(A2:A6,"=*West",B2:B6)"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A6", a2a6),
                new StringEval("=*West"),
                EvalFactory.CreateAreaEval("B2:B6", b2b6),
            };
            Confirm(16733.5, args);

            // "=AVERAGEIF(A2:A6,"<>*(New Office)",B2:B6)"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A6", a2a6),
                new StringEval("<>*(New Office)"),
                EvalFactory.CreateAreaEval("B2:B6", b2b6),

            };
            Confirm(18589, args);
        }
    }
}
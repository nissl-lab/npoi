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

namespace TestCases.SS.Formula.Functions {
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;

    /**
     * Test cases for AVERAGEIFS()
     *
     * @author Denilson Rodrigues
     */
    [TestFixture]
    public class TestAverageifs
    {
        private static OperationEvaluationContext EC = new OperationEvaluationContext(null, null, 0, 1, 0, null);

        private static ValueEval InvokeAverageifs(ValueEval[] args, OperationEvaluationContext ec)
        {
            return new AverageIfs().Evaluate(args, EC);
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
            ConfirmDouble(expectedResult, InvokeAverageifs(args, EC));
        }

        /**
         *  Example 1 from
         *  https://support.office.com/en-us/article/averageifs-function-48910c45-1fc0-4389-a028-f7c5c3001690
         */
        [Test]
        public void TestExample1()


        {
            // mimic test sample from https://support.office.com/en-us/article/averageifs-function-48910c45-1fc0-4389-a028-f7c5c3001690

            ValueEval[] b2b5 = new ValueEval[]
            {
                new StringEval("Quiz"),
                new StringEval("Grade"),
                new NumberEval(75),
                new NumberEval(94),
            };

            ValueEval[] c2c5 = new ValueEval[]
            {
                new StringEval("Quiz"),
                new StringEval("Grade"),
                new NumberEval(85),
                new NumberEval(80),
            };

            ValueEval[] d2d5 = new ValueEval[]
            {
                new StringEval("Exam"),
                new StringEval("Grade"),
                new NumberEval(87),
                new NumberEval(88),
            };

            // "=AVERAGEIFS(B2:B5, B2:B5, ">70", B2:B5, "<90")"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("B2:B5", b2b5),
                EvalFactory.CreateAreaEval("B2:B5", b2b5),
                new StringEval(">70"),
                EvalFactory.CreateAreaEval("B2:B5", b2b5),
                new StringEval("<90")
            };
            Confirm(75, args);

            // "=AVERAGEIFS(C2:C5, C2:C5, ">95")"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("C2:C5", c2c5),
                EvalFactory.CreateAreaEval("C2:C5", c2c5),
                new StringEval(">95"),
            };
            Assert.AreEqual(ErrorEval.VALUE_INVALID, InvokeAverageifs(args, EC));

            // "=AVERAGEIFS(D2:D5, D2:D5, "<>Incomplete", D2:D5, ">80")"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("D2:D5", d2d5),
                EvalFactory.CreateAreaEval("D2:D5", d2d5),
                new StringEval("<>Imcomplete"),
                EvalFactory.CreateAreaEval("D2:D5", d2d5),
                new StringEval(">80"),
            };
            Confirm(87.5, args);
        }

        /**
         *  Example 2 from
         *  https://support.office.com/en-us/article/averageifs-function-48910c45-1fc0-4389-a028-f7c5c3001690
         */
        [Test]
        public void TestExample2()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            ValueEval[] b2b7 = new ValueEval[]
            {
                new StringEval(""),
                new NumberEval(230000),
                new NumberEval(197000),
                new NumberEval(345678),
                new NumberEval(321900),
                new NumberEval(450000),
            };

            ValueEval[] c2c7 = new ValueEval[]
            {
                new StringEval(""),
                new StringEval("Issaquah"),
                new StringEval("Bellevue"),
                new StringEval("Bellevue"),
                new StringEval("Issaquah"),
                new StringEval("Bellevue"),
            };

            ValueEval[] d2d7 = new ValueEval[]
            {
                new StringEval(""),
                new NumberEval(3),
                new NumberEval(2),
                new NumberEval(4),
                new NumberEval(2),
                new NumberEval(5),
            };

            ValueEval[] e2e7 = new ValueEval[]
            {
                new StringEval(""),
                new StringEval("No"),
                new StringEval("Yes"),
                new StringEval("Yes"),
                new StringEval("Yes"),
                new StringEval("Yes"),
            };

            // "= AVERAGEIFS(B2: B7, C2: C7, "Bellevue", D2: D7, ">2", E2: E7, "Yes")"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                EvalFactory.CreateAreaEval("C2:C7", c2c7),
                new StringEval("Bellevue"),
                EvalFactory.CreateAreaEval("D2:D7", d2d7),
                new StringEval(">2"),
                EvalFactory.CreateAreaEval("E2:E7", e2e7),
                new StringEval("Yes"),
            };
            Confirm(397839, args);

            // "= AVERAGEIFS(B2: B7, C2: C7, "Issaquah", D2: D7, "<=3", E2: E7, "No")"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("B2:B7", b2b7),
                EvalFactory.CreateAreaEval("C2:C7", c2c7),
                new StringEval("Issaquah"),
                EvalFactory.CreateAreaEval("D2:D7", d2d7),
                new StringEval("<=3"),
                EvalFactory.CreateAreaEval("E2:E7", e2e7),
                new StringEval("No"),
            };
            Confirm(230000, args);
        }
    }
}
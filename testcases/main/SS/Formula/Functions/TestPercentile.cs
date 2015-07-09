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

    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;

    using NUnit.Framework;
    using TestCases.SS.Formula.Functions;

    /**
     * Testcase for Excel function PERCENTILE()
     *
     * @author T. Gordon
     */
    [TestFixture]
    public class TestPercentile
    {

        //[Test]
        //public void TestPercentile()
        //{
        //    testBasic();
        //    testUnusualArgs();
        //    testUnusualArgs2();
        //    testUnusualArgs3();
        //    testErrors();
        //    testErrors2();
        //}

        private static ValueEval invokePercentile(ValueEval[] args, ValueEval percentile)
        {
            AreaEval aeA = EvalFactory.CreateAreaEval("A1:A" + args.Length, args);
            ValueEval[] args2 = { aeA, percentile };
            return AggregateFunction.PERCENTILE.Evaluate(args2, -1, -1);
        }

        private void ConfirmPercentile(ValueEval percentile, ValueEval[] args, double expected)
        {
            ValueEval result = invokePercentile(args, percentile);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            double delta = 0.00000001;
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, delta);
        }

        private void ConfirmPercentile(ValueEval percentile, ValueEval[] args, ErrorEval expectedError)
        {
            ValueEval result = invokePercentile(args, percentile);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(expectedError.ErrorCode, ((ErrorEval)result).ErrorCode);
        }

        [Test]
        public void TestBasic()
        {
            ValueEval[] values = { new NumberEval(210.128), new NumberEval(65.2182), new NumberEval(32.231),
                new NumberEval(12.123), new NumberEval(45.32) };
            ValueEval percentile = new NumberEval(0.95);
            ConfirmPercentile(percentile, values, 181.14604);
        }

        [Test]
        public void TestBlanks()
        {
            ValueEval[] values = { new NumberEval(210.128), new NumberEval(65.2182), new NumberEval(32.231),
                BlankEval.instance, new NumberEval(45.32) };
            ValueEval percentile = new NumberEval(0.95);
            ConfirmPercentile(percentile, values, 188.39153);
        }

        [Test]
        public void TestUnusualArgs()
        {
            ValueEval[] values = { new NumberEval(1), new NumberEval(2), BoolEval.TRUE, BoolEval.FALSE };
            ValueEval percentile = new NumberEval(0.95);
            ConfirmPercentile(percentile, values, 1.95);
        }

        //percentile has to be between 0 and 1 - here we test less than zero
        [Test]
        public void TestUnusualArgs2()
        {
            ValueEval[] values = { new NumberEval(1), new NumberEval(2), };
            ValueEval percentile = new NumberEval(-0.1);
            ConfirmPercentile(percentile, values, ErrorEval.NUM_ERROR);
        }

        //percentile has to be between 0 and 1 - here we test more than 1
        [Test]
        public void TestUnusualArgs3()
        {
            ValueEval[] values = { new NumberEval(1), new NumberEval(2) };
            ValueEval percentile = new NumberEval(1.1);
            ConfirmPercentile(percentile, values, ErrorEval.NUM_ERROR);
        }

        //here we test where there are errors as part of inputs
        [Test]
        public void TestErrors()
        {
            ValueEval[] values = { new NumberEval(1), ErrorEval.NAME_INVALID, new NumberEval(3), ErrorEval.DIV_ZERO, };
            ValueEval percentile = new NumberEval(0.95);
            ConfirmPercentile(percentile, values, ErrorEval.NAME_INVALID);
        }

        //here we test where there are errors as part of inputs
        [Test]
        public void TestErrors2()
        {
            ValueEval[] values = { new NumberEval(1), new NumberEval(2), new NumberEval(3), ErrorEval.DIV_ZERO, };
            ValueEval percentile = new NumberEval(0.95);
            ConfirmPercentile(percentile, values, ErrorEval.DIV_ZERO);
        }
    }

}
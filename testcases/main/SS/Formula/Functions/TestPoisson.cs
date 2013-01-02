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
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;

    /**
     * Tests for Excel function POISSON(x,mean,cumulative)
     * @author Kalpesh Parmar
     */
    [TestFixture]
    public class TestPoisson
    {

        private static double DELTA = 1E-15;

        private static ValueEval invokePoisson(double x, double mean, bool cumulative)
        {

            ValueEval[] valueEvals = new ValueEval[3];
            valueEvals[0] = new NumberEval(x);
            valueEvals[1] = new NumberEval(mean);
            valueEvals[2] = BoolEval.ValueOf(cumulative);

            return NumericFunction.POISSON.Evaluate(valueEvals, -1, -1);
        }
        [Test]
        public void TestCumulativeProbability()
        {
            double x = 1;
            double mean = 0.2;
            double result = 0.9824769036935787; // known result

            NumberEval myResult = (NumberEval)invokePoisson(x, mean, true);

            Assert.AreEqual(myResult.NumberValue, result, DELTA);
        }
        [Test]
        public void TestNonCumulativeProbability()
        {
            double x = 0;
            double mean = 0.2;
            double result = 0.8187307530779818; // known result

            NumberEval myResult = (NumberEval)invokePoisson(x, mean, false);

            Assert.AreEqual(myResult.NumberValue, result, DELTA);
        }
        [Test]
        public void TestNegativeMean()
        {
            double x = 0;
            double mean = -0.2;

            ErrorEval myResult = (ErrorEval)invokePoisson(x, mean, false);

            Assert.AreEqual(ErrorEval.NUM_ERROR.ErrorCode, myResult.ErrorCode);
        }
        [Test]
        public void TestNegativeX()
        {
            double x = -1;
            double mean = 0.2;

            ErrorEval myResult = (ErrorEval)invokePoisson(x, mean, false);

            Assert.AreEqual(ErrorEval.NUM_ERROR.ErrorCode, myResult.ErrorCode);
        }


        [Test]
        public void TestXAsDecimalNumber()
        {
            double x = 1.1;
            double mean = 0.2;
            double result = 0.9824769036935787; // known result

            NumberEval myResult = (NumberEval)invokePoisson(x, mean, true);

            Assert.AreEqual(myResult.NumberValue, result, DELTA);
        }
        [Test]
        public void TestXZeroMeanZero()
        {
            double x = 0;
            double mean = 0;
            double result = 1; // known result in excel

            NumberEval myResult = (NumberEval)invokePoisson(x, mean, true);

            Assert.AreEqual(myResult.NumberValue, result, DELTA);
        }
    }

}
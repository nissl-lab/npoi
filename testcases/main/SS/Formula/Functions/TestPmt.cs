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
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;

    /**
     * @author Josh Micich
     */
    [TestFixture]
    public class TestPmt
    {

        private static void Confirm(double expected, NumberEval ne)
        {
            // only asserting accuracy to 4 fractional digits
            Assert.AreEqual(expected, ne.NumberValue, 0.00005);
        }
        private static ValueEval invoke(ValueEval[] args)
        {
            return FinanceFunction.PMT.Evaluate(args, -1, (short)-1);
        }
        /**
         * Invocation when not expecting an error result
         */
        private static NumberEval invokeNormal(ValueEval[] args)
        {
            ValueEval ev = invoke(args);
            if (ev is ErrorEval)
            {
                Assert.Fail("Normal Evaluation failed with error code: "
                        + ev.ToString());
            }
            return (NumberEval)ev;
        }

        private static void Confirm(double expected, double rate, double nper, double pv, double fv, bool isBeginning)
        {
            ValueEval[] args = {
				new NumberEval(rate),
				new NumberEval(nper),
				new NumberEval(pv),
				new NumberEval(fv),
				new NumberEval(isBeginning ? 1 : 0),
		};
            Confirm(expected, invokeNormal(args));
        }

        [Test]
        public void TestBasic()
        {
            Confirm(-1037.0321, (0.08 / 12), 10, 10000, 0, false);
            Confirm(-1030.1643, (0.08 / 12), 10, 10000, 0, true);
        }
        [Test]
        public void Test3args()
        {

            ValueEval[] args = {
				new NumberEval(0.005),
				new NumberEval(24),
				new NumberEval(1000),
		};
            ValueEval ev = invoke(args);
            if (ev is ErrorEval)
            {
                ErrorEval err = (ErrorEval)ev;
                if (err.ErrorCode == FormulaError.VALUE.Code)
                {
                    Assert.Fail("Identified bug 44691");
                }
            }

            Confirm(-44.3206, invokeNormal(args));
        }
    }

}
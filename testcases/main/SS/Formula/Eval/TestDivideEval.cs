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

    using NUnit.Framework;
    using NPOI.SS.Formula.Eval;
    using TestCases.SS.Formula.Functions;

    /**
     * Test for divide operator Evaluator.
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestDivideEval
    {

        private static void Confirm(ValueEval arg0, ValueEval arg1, double expectedResult)
        {
            ValueEval[] args = {
			arg0, arg1,
		};

            double result = NumericFunctionInvoker.Invoke(EvalInstances.Divide, args, 0, 0);

            Assert.AreEqual(expectedResult, result, 0);
        }
        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US"); 
            
            Confirm(new NumberEval(5), new NumberEval(2), 2.5);
            Confirm(new NumberEval(3), new NumberEval(16), 0.1875);
            Confirm(new NumberEval(-150), new NumberEval(-15), 10.0);
            Confirm(new StringEval("0.2"), new NumberEval(0.05), 4.0);
            Confirm(BoolEval.TRUE, new StringEval("-0.2"), -5.0);
        }
        [Test]
        public void Test1x1Area()
        {
            AreaEval ae0 = EvalFactory.CreateAreaEval("B2:B2", new ValueEval[] { new NumberEval(50), });
            AreaEval ae1 = EvalFactory.CreateAreaEval("C2:C2", new ValueEval[] { new NumberEval(10), });
            Confirm(ae0, ae1, 5);
        }
        [Test]
        public void TestDivZero()
        {
            ValueEval[] args = {
			new NumberEval(5), NumberEval.ZERO,
		};
            ValueEval result = EvalInstances.Divide.Evaluate(args, 0, (short)0);
            Assert.AreEqual(ErrorEval.DIV_ZERO, result);
        }
    }

}
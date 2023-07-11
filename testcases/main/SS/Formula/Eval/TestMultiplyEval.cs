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
     *  Test for multiply operator Evaluator.
     *  
     *  
     */
    [TestFixture]
    public class TestMultiplyEval
    {

        private static void Confirm(ValueEval arg0, ValueEval arg1, double expectedResult)
        {
            ValueEval[] args = {
			    arg0, arg1,
		    };

            double result = NumericFunctionInvoker.Invoke(EvalInstances.Multiply, args, 0, 0);

            Assert.AreEqual(expectedResult, result, 0);
        }
        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            // issue #1055, use decimal to handle precision issue like (20000 * 0.000027 = 0.539999 in double)
            Confirm(new NumberEval(20000), new NumberEval(0.000027), 0.54);            
        }
        
    }

}
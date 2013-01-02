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

using NPOI.SS.Formula.Eval;
using NUnit.Framework;
using NPOI.SS.Formula.Functions;
namespace TestCases.SS.Formula.Functions
{

    /**
     * Tests for Excel function AVERAGE()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestAverage
    {

        private static ValueEval InvokeAverage(ValueEval[] args)
        {
            return AggregateFunction.AVERAGE.Evaluate(args, -1, (short)-1);
        }

        private void ConfirmAverage(ValueEval[] args, double expected)
        {
            ValueEval result = InvokeAverage(args);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0);
        }

        private void ConfirmAverage(ValueEval[] args, ErrorEval expectedError)
        {
            ValueEval result = InvokeAverage(args);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(expectedError.ErrorCode, ((ErrorEval)result).ErrorCode);
        }
        [Test]
        public void TestBasic()
        {

            ValueEval[] values = {
				new NumberEval(1),
				new NumberEval(2),
				new NumberEval(3),
				new NumberEval(4),
		};

            ConfirmAverage(values, 2.5);

            values = new ValueEval[] {
				new NumberEval(1),
				new NumberEval(2),
				BlankEval.instance,
				new NumberEval(3),
				BlankEval.instance,
				new NumberEval(4),
				BlankEval.instance,
		};

            ConfirmAverage(values, 2.5);
        }

        /**
         * Valid cases where values are not pure numbers
         */
        [Test]
        public void TestUnusualArgs()
        {
            ValueEval[] values = {
				new NumberEval(1),
				new NumberEval(2),
				BoolEval.TRUE,
				BoolEval.FALSE,
		};

            ConfirmAverage(values, 1.0);

        }
        [Test]
        public void TestErrors()
        {
            ValueEval[] values = {
				new NumberEval(1),
				ErrorEval.NAME_INVALID,
				new NumberEval(3),
				ErrorEval.DIV_ZERO,
		};
            ConfirmAverage(values, ErrorEval.NAME_INVALID);
        }
    }

}
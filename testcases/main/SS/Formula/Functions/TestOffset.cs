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

namespace NPOI.SS.Formula.functions;

using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.SS.Formula.Eval.EvaluationException;
using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.functions.Offset.LinearOffsetRange;

/**
 * Tests for OFFSET function implementation
 *
 * @author Josh Micich
 */
public class TestOffset  {

	private static void ConfirmDoubleConvert(double doubleVal, int expected) {
		try {
			Assert.AreEqual(expected, Offset.EvaluateIntArg(new NumberEval(doubleVal), -1, -1));
		} catch (EvaluationException e) {
			throw new AssertionFailedError("Unexpected error '" + e.GetErrorEval().ToString() + "'.");
		}
	}
	/**
	 * Excel's double to int conversion (for function 'OFFSET()') behaves more like Math.floor().
	 * Note - negative values are not symmetrical
	 * Fractional values are silently tRuncated.
	 * TRuncation is toward negative infInity.
	 */
	public void TestDoubleConversion() {

		ConfirmDoubleConvert(100.09, 100);
		ConfirmDoubleConvert(100.01, 100);
		ConfirmDoubleConvert(100.00, 100);
		ConfirmDoubleConvert(99.99, 99);

		ConfirmDoubleConvert(+2.01, +2);
		ConfirmDoubleConvert(+2.00, +2);
		ConfirmDoubleConvert(+1.99, +1);
		ConfirmDoubleConvert(+1.01, +1);
		ConfirmDoubleConvert(+1.00, +1);
		ConfirmDoubleConvert(+0.99,  0);
		ConfirmDoubleConvert(+0.01,  0);
		ConfirmDoubleConvert( 0.00,  0);
		ConfirmDoubleConvert(-0.01, -1);
		ConfirmDoubleConvert(-0.99, -1);
		ConfirmDoubleConvert(-1.00, -1);
		ConfirmDoubleConvert(-1.01, -2);
		ConfirmDoubleConvert(-1.99, -2);
		ConfirmDoubleConvert(-2.00, -2);
		ConfirmDoubleConvert(-2.01, -3);
	}

	public void TestLinearOffsetRange() {
		LinearOffsetRange lor;

		lor = new LinearOffsetRange(3, 2);
		Assert.AreEqual(3, lor.GetFirstIndex());
		Assert.AreEqual(4, lor.GetLastIndex());
		lor = lor.normaliseAndTranslate(0); // expected no change
		Assert.AreEqual(3, lor.GetFirstIndex());
		Assert.AreEqual(4, lor.GetLastIndex());

		lor = lor.normaliseAndTranslate(5);
		Assert.AreEqual(8, lor.GetFirstIndex());
		Assert.AreEqual(9, lor.GetLastIndex());

		// negative length

		lor = new LinearOffsetRange(6, -4).normaliseAndTranslate(0);
		Assert.AreEqual(3, lor.GetFirstIndex());
		Assert.AreEqual(6, lor.GetLastIndex());


		// bounds Checking
		lor = new LinearOffsetRange(0, 100);
		Assert.IsFalse(lor.IsOutOfBounds(0, 16383));
		lor = lor.normaliseAndTranslate(16300);
		Assert.IsTrue(lor.IsOutOfBounds(0, 16383));
		Assert.IsFalse(lor.IsOutOfBounds(0, 65535));
	}
}


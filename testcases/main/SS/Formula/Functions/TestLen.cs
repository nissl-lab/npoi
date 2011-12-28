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

using junit.framework.TestCase;

using NPOI.SS.Formula.Eval.BlankEval;
using NPOI.SS.Formula.Eval.BoolEval;
using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.Eval.StringEval;
using NPOI.SS.Formula.Eval.ValueEval;
/**
 * Tests for Excel function LEN()
 *
 * @author Josh Micich
 */
public class TestLen  {

	private static ValueEval invokeLen(ValueEval text) {
		ValueEval[] args = new ValueEval[] { text, };
		return TextFunction.LEN.Evaluate(args, -1, (short)-1);
	}

	private void ConfirmLen(ValueEval text, int expected) {
		ValueEval result = invokeLen(text);
		Assert.AreEqual(NumberEval.class, result.GetType());
		Assert.AreEqual(expected, ((NumberEval)result).GetNumberValue(), 0);
	}

	private void ConfirmLen(ValueEval text, ErrorEval expectedError) {
		ValueEval result = invokeLen(text);
		Assert.AreEqual(ErrorEval.class, result.GetType());
		Assert.AreEqual(expectedError.GetErrorCode(), ((ErrorEval)result).GetErrorCode());
	}

	public void TestBasic() {

		ConfirmLen(new StringEval("galactic"), 8);
	}

	/**
	 * Valid cases where text arg is not exactly a string
	 */
	public void TestUnusualArgs() {

		// text (first) arg type is number, other args are strings with fractional digits
		ConfirmLen(new NumberEval(123456), 6);
		ConfirmLen(BoolEval.FALSE, 5);
		ConfirmLen(BoolEval.TRUE, 4);
		ConfirmLen(BlankEval.instance, 0);
	}

	public void TestErrors() {
		ConfirmLen(ErrorEval.NAME_INVALID, ErrorEval.NAME_INVALID);
	}
}


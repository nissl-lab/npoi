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

using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.Formula.Eval.ValueEval;
using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.Eval.StringEval;

/**
 * Test case for TRUNC()
 *
 * @author Stephen Wolke (smwolke at geistig.com)
 */
public class TestTRunc : AbstractNumericTestCase {
	private static NumericFunction F = null;
	public void TestTRuncWithStringArg() {

		ValueEval strArg = new StringEval("abc");
		ValueEval[] args = { strArg, new NumberEval(2) };
		ValueEval result = F.TRUNC.Evaluate(args, -1, (short)-1);
		Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
	}

	public void TestTRuncWithWholeNumber() {
		ValueEval[] args = { new NumberEval(200), new NumberEval(2) };
		@SuppressWarnings("static-access")
		ValueEval result = F.TRUNC.Evaluate(args, -1, (short)-1);
		Assert.AreEqual("TRUNC", (new NumberEval(200d)).GetNumberValue(), ((NumberEval)result).GetNumberValue());
	}
	
	public void TestTRuncWithDecimalNumber() {
		ValueEval[] args = { new NumberEval(2.612777), new NumberEval(3) };
		@SuppressWarnings("static-access")
		ValueEval result = F.TRUNC.Evaluate(args, -1, (short)-1);
		Assert.AreEqual("TRUNC", (new NumberEval(2.612d)).GetNumberValue(), ((NumberEval)result).GetNumberValue());
	}
	
	public void TestTRuncWithDecimalNumberOneArg() {
		ValueEval[] args = { new NumberEval(2.612777) };
		ValueEval result = F.TRUNC.Evaluate(args, -1, (short)-1);
		Assert.AreEqual("TRUNC", (new NumberEval(2d)).GetNumberValue(), ((NumberEval)result).GetNumberValue());
	}
}


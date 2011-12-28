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

using NPOI.SS.Formula.Eval.AreaEval;
using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.Eval.NumericValueEval;
using NPOI.SS.Formula.Eval.StringEval;
using NPOI.SS.Formula.Eval.ValueEval;

/**
 * Test cases for SUMPRODUCT()
 *
 * @author Josh Micich
 */
public class TestSumif  {
	private static NumberEval _30 = new NumberEval(30);
	private static NumberEval _40 = new NumberEval(40);
	private static NumberEval _50 = new NumberEval(50);
	private static NumberEval _60 = new NumberEval(60);

	private static ValueEval invokeSumif(int rowIx, int colIx, ValueEval...args) {
		return new Sumif().Evaluate(args, rowIx, colIx);
	}
	private static void ConfirmDouble(double expected, ValueEval actualEval) {
		if(!(actualEval is NumericValueEval)) {
			throw new AssertionFailedError("Expected numeric result");
		}
		NumericValueEval nve = (NumericValueEval)actualEval;
		Assert.AreEqual(expected, nve.GetNumberValue(), 0);
	}

	public void TestBasic() {
		ValueEval[] arg0values = new ValueEval[] { _30, _30, _40, _40, _50, _50  };
		ValueEval[] arg2values = new ValueEval[] { _30, _40, _50, _60, _60, _60 };

		AreaEval arg0;
		AreaEval arg2;

		arg0 = EvalFactory.CreateAreaEval("A3:B5", arg0values);
		arg2 = EvalFactory.CreateAreaEval("D1:E3", arg2values);

		Confirm(60.0, arg0, new NumberEval(30.0));
		Confirm(70.0, arg0, new NumberEval(30.0), arg2);
		Confirm(100.0, arg0, new StringEval(">45"));

	}
	private static void Confirm(double expectedResult, ValueEval...args) {
		ConfirmDouble(expectedResult, invokeSumif(-1, -1, args));
	}


	/**
	 * Test for bug observed near svn r882931
	 */
	public void TestCriteriaArgRange() {
		ValueEval[] arg0values = new ValueEval[] { _50, _60, _50, _50, _50, _30,  };
		ValueEval[] arg1values = new ValueEval[] { _30, _40, _50, _60,  };

		AreaEval arg0;
		AreaEval arg1;
		ValueEval ve;

		arg0 = EvalFactory.CreateAreaEval("A3:B5", arg0values);
		arg1 = EvalFactory.CreateAreaEval("A2:D2", arg1values); // single row range

		ve = invokeSumif(0, 2, arg0, arg1);  // invoking from cell C1
		if (ve is NumberEval) {
			NumberEval ne = (NumberEval) ve;
			if (ne.GetNumberValue() == 30.0) {
				throw new AssertionFailedError("identified error in SUMIF - criteria arg not Evaluated properly");
			}
		}

		ConfirmDouble(200, ve);

		arg0 = EvalFactory.CreateAreaEval("C1:D3", arg0values);
		arg1 = EvalFactory.CreateAreaEval("B1:B4", arg1values); // single column range

		ve = invokeSumif(3, 0, arg0, arg1); // invoking from cell A4

		ConfirmDouble(60, ve);
	}
}


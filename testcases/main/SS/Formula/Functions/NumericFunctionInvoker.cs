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

using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.Formula.Eval.NumericValueEval;
using NPOI.SS.Formula.Eval.ValueEval;
using NPOI.SS.Formula.Eval.NotImplementedException;

/**
 * Test helper class for invoking functions with numeric results.
 *
 * @author Josh Micich
 */
public class NumericFunctionInvoker {

	private NumericFunctionInvoker() {
		// no instances of this class
	}

	private static class NumericEvalEx : Exception {
		public NumericEvalEx(String msg) {
			base(msg);
		}
	}

	/**
	 * Invokes the specified function with the arguments.
	 * <p/>
	 * Assumes that the cell coordinate parameters of
	 *  <code>Function.Evaluate(args, srcCellRow, srcCellCol)</code>
	 * are not required.
	 * <p/>
	 * This method cannot be used for Confirming error return codes.  Any non-numeric Evaluation
	 * result causes the current junit Test to fail.
	 */
	public static double invoke(Function f, ValueEval[] args) {
		return invoke(f, args, -1, -1);
	}
	/**
	 * Invokes the specified operator with the arguments.
	 * <p/>
	 * This method cannot be used for Confirming error return codes.  Any non-numeric Evaluation
	 * result causes the current junit Test to fail.
	 */
	public static double invoke(Function f, ValueEval[] args, int srcCellRow, int srcCellCol) {
		try {
			return invokeInternal(f, args, srcCellRow, srcCellCol);
		} catch (NumericEvalEx e) {
			throw new AssertionFailedError("Evaluation of function (" + f.GetType().GetName()
					+ ") failed: " + e.GetMessage());
		}
	}
	/**
	 * Formats nicer error messages for the junit output
	 */
	private static double invokeInternal(Function target, ValueEval[] args, int srcCellRow, int srcCellCol)
				throws NumericEvalEx {
		ValueEval EvalResult;
		try {
			EvalResult = target.Evaluate(args, srcCellRow, (short)srcCellCol);
		} catch (NotImplementedException e) {
			throw new NumericEvalEx("Not implemented:" + e.GetMessage());
		}

		if(EvalResult == null) {
			throw new NumericEvalEx("Result object was null");
		}
		if(EvalResult is ErrorEval) {
			ErrorEval ee = (ErrorEval) EvalResult;
			throw new NumericEvalEx(formatErrorMessage(ee));
		}
		if(!(EvalResult is NumericValueEval)) {
			throw new NumericEvalEx("Result object type (" + EvalResult.GetType().GetName()
					+ ") is invalid.  Expected implementor of ("
					+ NumericValueEval.class.GetName() + ")");
		}

		NumericValueEval result = (NumericValueEval) EvalResult;
		return result.GetNumberValue();
	}
	private static String formatErrorMessage(ErrorEval ee) {
		if(errorCodesAreEqual(ee, ErrorEval.VALUE_INVALID)) {
			return "Error code: #VALUE! (invalid value)";
		}
		return "Error code=" + ee.GetErrorCode();
	}
	private static bool errorCodesAreEqual(ErrorEval a, ErrorEval b) {
		if(a==b) {
			return true;
		}
		return a.GetErrorCode() == b.GetErrorCode();
	}
}


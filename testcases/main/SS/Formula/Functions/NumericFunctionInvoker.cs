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
    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;

    /**
     * Test helper class for invoking functions with numeric results.
     *
     * @author Josh Micich
     */
    public class NumericFunctionInvoker
    {

        private NumericFunctionInvoker()
        {
            // no instances of this class
        }

        private class NumericEvalEx : Exception
        {
            public NumericEvalEx(String msg)
                : base(msg)
            {
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
        public static double Invoke(Function f, ValueEval[] args)
        {
            return Invoke(f, args, -1, -1);
        }
        /**
         * Invokes the specified operator with the arguments.
         * <p/>
         * This method cannot be used for Confirming error return codes.  Any non-numeric Evaluation
         * result causes the current junit Test to fail.
         */
        public static double Invoke(Function f, ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            try
            {
                return invokeInternal(f, args, srcCellRow, srcCellCol);
            }
            catch (NumericEvalEx e)
            {
                throw new AssertionException("Evaluation of function (" + f.GetType().Name
                        + ") failed: " + e.Message);
            }
        }
        /**
         * Formats nicer error messages for the junit output
         */
        private static double invokeInternal(Function target, ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            ValueEval EvalResult = null;
            try
            {
                EvalResult = target.Evaluate(args, srcCellRow, (short)srcCellCol);
            }
            catch (NotImplementedException e)
            {
                throw new NumericEvalEx("Not implemented:" + e.Message);
            }

            if (EvalResult == null)
            {
                throw new NumericEvalEx("Result object was null");
            }
            if (EvalResult is ErrorEval)
            {
                ErrorEval ee = (ErrorEval)EvalResult;
                throw new NumericEvalEx(formatErrorMessage(ee));
            }
            if (!(EvalResult is NumericValueEval))
            {
                throw new NumericEvalEx("Result object type (" + EvalResult.GetType().Name
                        + ") is invalid.  Expected implementor of ("
                        + typeof(NumericValueEval).Name + ")");
            }

            NumericValueEval result = (NumericValueEval)EvalResult;
            return result.NumberValue;
        }
        private static String formatErrorMessage(ErrorEval ee)
        {
            if (errorCodesAreEqual(ee, ErrorEval.VALUE_INVALID))
            {
                return "Error code: #VALUE! (invalid value)";
            }
            return "Error code=" + ee.ErrorCode;
        }
        private static bool errorCodesAreEqual(ErrorEval a, ErrorEval b)
        {
            if (a == b)
            {
                return true;
            }
            return a.ErrorCode == b.ErrorCode;
        }
    }
}

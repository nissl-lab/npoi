/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace TestCases.HSSF.Record.Formula.Functions
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.HSSF.Record.Formula.Functions;

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
         *  <c>Function.evaluate(args, srcCellRow, srcCellCol)</c>
         * are not required.
         * <p/>
         * This method cannot be used for Confirming error return codes.  Any non-numeric evaluation
         * result causes the current junit test to fail.
         */
        public static double Invoke(Function f, ValueEval[] args)
        {
            try
            {
                return InvokeInternal(f, args, -1, -1);
            }
            catch (NumericEvalEx e)
            {
                throw new AssertFailedException("Evaluation of function (" + f.GetType().Name
                        + ") failed: " + e.Message);
            }

        }
        /**
         * Invokes the specified operator with the arguments.
         * <p/>
         * This method cannot be used for Confirming error return codes.  Any non-numeric evaluation
         * result causes the current junit test to fail.
         */
        public static double Invoke(Function f, ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            try
            {
                return InvokeInternal(f, args, srcCellRow, srcCellCol);
            }
            catch (NumericEvalEx e)
            {
                throw new AssertFailedException("Evaluation of function (" + f.GetType().Name
                        + ") failed: " + e.Message);
            }

        }
        /**
         * Formats nicer error messages for the junit output
         */
        private static double InvokeInternal(Function target, ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            ValueEval evalResult;
            try
            {
                evalResult = target.Evaluate(args, srcCellRow,srcCellCol);
            }
            catch (NotImplementedException e)
            {
                throw new NumericEvalEx("Not implemented:" + e.Message);
            }

            if (evalResult == null)
            {
                throw new NumericEvalEx("Result object was null");
            }
            if (evalResult is ErrorEval)
            {
                ErrorEval ee = (ErrorEval)evalResult;
                throw new NumericEvalEx(FormatErrorMessage(ee));
            }
            if (!(evalResult is NumericValueEval))
            {
                throw new NumericEvalEx("Result object type (" + evalResult.GetType().Name
                        + ") is1 invalid.  Expected implementor of ("
                        + typeof(NumericValueEval).Name + ")");
            }

            NumericValueEval result = (NumericValueEval)evalResult;
            return result.NumberValue;
        }
        private static String FormatErrorMessage(ErrorEval ee)
        {
            if (ErrorCodesAreEqual(ee, ErrorEval.FUNCTION_NOT_IMPLEMENTED))
            {
                return "Function not implemented";
            }
            if (ErrorCodesAreEqual(ee, ErrorEval.VALUE_INVALID))
            {
                return "Error code: #VALUE! (invalid value)";
            }
            return "Error code=" + ee.ErrorCode;
        }
        private static bool ErrorCodesAreEqual(ErrorEval a, ErrorEval b)
        {
            if (a == b)
            {
                return true;
            }
            return a.ErrorCode == b.ErrorCode;
        }

    }
}
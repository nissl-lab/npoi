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
using System;
using NPOI.SS.Util;
namespace NPOI.SS.Formula.Functions
{



    /**
     * Implementation for Excel DELTA() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>DELTA </b>(<b>number1</b>,<b>number2</b> )<br/>
     * <p/>
     * Tests whether two values are Equal. Returns 1 if number1 = number2; returns 0 otherwise.
     * Use this function to filter a Set of values. For example, by summing several DELTA functions
     * you calculate the count of equal pairs. This function is also known as the Kronecker Delta function.
     *
     * <ul>
     *     <li>If number1 is nonnumeric, DELTA returns the #VALUE! error value.</li>
     *     <li>If number2 is nonnumeric, DELTA returns the #VALUE! error value.</li>
     * </ul>
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Delta : Fixed2ArgFunction, FreeRefFunction
    {
        public static FreeRefFunction instance = new Delta();
        private static NumberEval ONE = new NumberEval(1);
        private static NumberEval ZERO = new NumberEval(0);


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1, ValueEval arg2)
        {
            ValueEval veText1;
            try
            {
                veText1 = OperandResolver.GetSingleValue(arg1, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            String strText1 = OperandResolver.CoerceValueToString(veText1);
            Double number1 = OperandResolver.ParseDouble(strText1);
            if (double.IsNaN(number1))
            {
                return ErrorEval.VALUE_INVALID;
            }

            ValueEval veText2;
            try
            {
                veText2 = OperandResolver.GetSingleValue(arg2, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            String strText2 = OperandResolver.CoerceValueToString(veText2);
            Double number2 = OperandResolver.ParseDouble(strText2);
            if (double.IsNaN(number2))
            {
                return ErrorEval.VALUE_INVALID;
            }

            //int result = new BigDecimal(number1).CompareTo(new BigDecimal(number2));
            int result = NumberComparer.Compare(number1, number2);
            return result == 0 ? ONE : ZERO;
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 2)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
            }

            return ErrorEval.VALUE_INVALID;
        }
    }
}
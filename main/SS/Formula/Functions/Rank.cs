/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list.

     * Syntax:
     *    RANK(number,ref,order)
     *       Number   is the number whose rank you want to find.
     *       Ref     is an array of, or a reference to, a list of numbers. Nonnumeric values in ref are ignored.
     *       Order   is a number specifying how to rank number.

     * If order is 0 (zero) or omitted, Microsoft Excel ranks number as if ref were a list sorted in descending order.
     * If order is any nonzero value, Microsoft Excel ranks number as if ref were a list sorted in ascending order.
     * 
     * @author Rubin Wang
     */
    public class Rank : Var2or3ArgFunction
    {

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {

            AreaEval aeRange;
            double result;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                result = OperandResolver.CoerceValueToDouble(ve);
                if (Double.IsNaN(result) || Double.IsInfinity(result))
                {
                    throw new EvaluationException(ErrorEval.NUM_ERROR);
                }
                aeRange = ConvertRangeArg(arg1);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return eval(srcRowIndex, srcColumnIndex, result, aeRange, true);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1, ValueEval arg2)
        {

            AreaEval aeRange;
            double result;
            bool order = false;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                result = OperandResolver.CoerceValueToDouble(ve);
                if (Double.IsNaN(result) || Double.IsInfinity(result))
                {
                    throw new EvaluationException(ErrorEval.NUM_ERROR);
                }
                aeRange = ConvertRangeArg(arg1);

                ve = OperandResolver.GetSingleValue(arg2, srcRowIndex, srcColumnIndex);
                int order_value = OperandResolver.CoerceValueToInt(ve);
                if (order_value == 0)
                {
                    order = true;
                }
                else if (order_value == 1)
                {
                    order = false;
                }
                else throw new EvaluationException(ErrorEval.NUM_ERROR);

            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return eval(srcRowIndex, srcColumnIndex, result, aeRange, order);
        }

        private static ValueEval eval(int srcRowIndex, int srcColumnIndex, double arg0, AreaEval aeRange, bool descending_order)
        {

            int rank = 1;
            int height = aeRange.Height;
            int width = aeRange.Width;
            for (int r = 0; r < height; r++)
            {
                for (int c = 0; c < width; c++)
                {

                    Double value = GetValue(aeRange, r, c);
                    if (Double.IsNaN(value)) continue;
                    if (descending_order && value > arg0 || !descending_order && value < arg0)
                    {
                        rank++;
                    }
                }
            }
            return new NumberEval(rank);
        }

        private static Double GetValue(AreaEval aeRange, int relRowIndex, int relColIndex)
        {

            ValueEval addend = aeRange.GetRelativeValue(relRowIndex, relColIndex);
            if (addend is NumberEval)
            {
                return ((NumberEval)addend).NumberValue;
            }
            // everything else (including string and boolean values) counts as zero
            return Double.NaN;
        }

        private static AreaEval ConvertRangeArg(ValueEval eval)
        {
            if (eval is AreaEval)
            {
                return (AreaEval)eval;
            }
            if (eval is RefEval)
            {
                return ((RefEval)eval).Offset(0, 0, 0, 0);
            }
            throw new EvaluationException(ErrorEval.VALUE_INVALID);
        }

    }

}
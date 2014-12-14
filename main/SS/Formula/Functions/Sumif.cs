/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;

    /**
     * Implementation for the Excel function SUMIF<p>
     *
     * Syntax : <br/>
     *  SUMIF ( <b>range</b>, <b>criteria</b>, sum_range ) <br/>
     *    <table border="0" cellpadding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>range</th><td>The range over which criteria is applied.  Also used for addend values when the third parameter is not present</td></tr>
     *      <tr><th>criteria</th><td>The value or expression used to filter rows from <b>range</b></td></tr>
     *      <tr><th>sum_range</th><td>Locates the top-left corner of the corresponding range of addends - values to be added (after being selected by the criteria)</td></tr>
     *    </table><br/>
     * </p>
     * @author Josh Micich
     */
    public class Sumif : Var2or3ArgFunction
    {
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            AreaEval aeRange;
            try
            {
                aeRange = ConvertRangeArg(arg0);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return Eval(srcRowIndex, srcColumnIndex, arg1, aeRange, aeRange);
        }

        public override  ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2)
        {
            AreaEval aeRange;
            AreaEval aeSum;
            try
            {
                aeRange = ConvertRangeArg(arg0);
                aeSum = CreateSumRange(arg2, aeRange);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return Eval(srcRowIndex, srcColumnIndex, arg1, aeRange, aeSum);
        }

        private static ValueEval Eval(int srcRowIndex, int srcColumnIndex, ValueEval arg1, AreaEval aeRange,
                AreaEval aeSum)
        {

            // TODO - junit to prove last arg must be srcColumnIndex and not srcRowIndex
            IMatchPredicate mp = Countif.CreateCriteriaPredicate(arg1, srcRowIndex, srcColumnIndex);
            double result = SumMatchingCells(aeRange, mp, aeSum);
            return new NumberEval(result);

        }

        private static double SumMatchingCells(AreaEval aeRange, IMatchPredicate mp, AreaEval aeSum)
        {
            int height = aeRange.Height;
            int width = aeRange.Width;

            double result = 0.0;

            for (int r = 0; r < height; r++)
            {
                for (int c = 0; c < width; c++)
                {
                    result += Accumulate(aeRange, mp, aeSum, r, c);
                }
            }
            return result;
        }



        private static double Accumulate(AreaEval aeRange, IMatchPredicate mp, AreaEval aeSum, int relRowIndex,
                int relColIndex)
        {
            if (!mp.Matches(aeRange.GetRelativeValue(relRowIndex, relColIndex)))
            {
                return 0.0;
            }

            ValueEval addend = aeSum.GetRelativeValue(relRowIndex, relColIndex);
            if (addend is NumberEval)
            {
                return ((NumberEval)addend).NumberValue;
            }
            // everything else (including string and boolean values) counts as zero
            return 0.0;
        }



        /**
         * @return a range of the same dimensions as aeRange using eval to define the top left corner.
         * @throws EvaluationException if eval is not a reference
         */

        private static AreaEval CreateSumRange(ValueEval eval, AreaEval aeRange)
        {
            if (eval is AreaEval)
            {
                return ((AreaEval)eval).Offset(0, aeRange.Height - 1, 0, aeRange.Width - 1);
            }
            if (eval is RefEval)
            {
                return ((RefEval)eval).Offset(0, aeRange.Height - 1, 0, aeRange.Width - 1);
            }
            throw new EvaluationException(ErrorEval.VALUE_INVALID);

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

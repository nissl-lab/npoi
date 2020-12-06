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
using NPOI.SS.Formula.Eval;
namespace NPOI.SS.Formula.Functions {
    /**
     * Implementation for the Excel function AverageIfs<br/>
     * <p>
     * Syntax : <br/>
     *  AverageIfs ( <b>average_range</b>, <b>criteria_range1</b>, <b>criteria1</b>,
     *  [<b>criteria_range2</b>,  <b>criteria2</b>], ...) <br/>
     *    <ul>
     *      <li><b>average_range</b> Required. One or more cells to get the average, including numbers or names, ranges,
     *      or cell references that contain numbers. Blank and text values are ignored.</li>
     *      <li><b>criteria1_range</b> Required. The first range in which
     *      to evaluate the associated criteria.</li>
     *      <li><b>criteria1</b> Required. The criteria in the form of a number, expression,
     *        cell reference, or text that define which cells in the criteria_range1
     *        argument will be counted</li>
     *      <li><b> criteria_range2, criteria2, ...</b>    Optional. Additional ranges and their associated criteria.
     *      Up to 127 range/criteria pairs are allowed.</li>
     *    </ul>
     * </p>
     *
     * @author Yegor Kozlov
     */
    public class AverageIfs : FreeRefFunction
    {
        public static FreeRefFunction instance = new AverageIfs();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length % 2 == 0)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                AreaEval avgRange = ConvertRangeArg(args[0]);

                // collect pairs of ranges and criteria
                AreaEval[] ae = new AreaEval[(args.Length - 1) / 2];
                IMatchPredicate[] mp = new IMatchPredicate[ae.Length];
                for (int i = 1, k = 0; i < args.Length; i += 2, k++)
                {
                    ae[k] = ConvertRangeArg(args[i]);
                    mp[k] = Countif.CreateCriteriaPredicate(args[i + 1], ec.RowIndex, ec.ColumnIndex);
                }

                ValidateCriteriaRanges(ae, avgRange);
                Sumifs.ValidateCriteria(mp);

                double result = GetAvgFromMatchingCells(ae, mp, avgRange);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        /**
         * Verify that each <code>criteriaRanges</code> argument contains the same number of rows and columns
         * as the <code>avgRange</code> argument
         *
         * @throws EvaluationException if
         */
        private void ValidateCriteriaRanges(AreaEval[] criteriaRanges, AreaEval avgRange)
        {
            foreach (AreaEval r in criteriaRanges)
            {
                if (r.Height != avgRange.Height ||
                   r.Width != avgRange.Width)
                {
                    throw EvaluationException.InvalidValue();
                }
            }
        }

        /**
         *
         * @param ranges  criteria ranges, each range must be of the same dimensions as <code>aeAvg</code>
         * @param predicates  array of predicates, a predicate for each value in <code>ranges</code>
         * @param aeAvg  the range to calculate
         *
         * @return the computed value
         */
        private static double GetAvgFromMatchingCells(AreaEval[] ranges, IMatchPredicate[] predicates, AreaEval aeAvg)
        {
            int height = aeAvg.Height;
            int width = aeAvg.Width;

            double sum = 0.0;
            int valuesCount = 0;
            for (int r = 0; r < height; r++)
            {
                for (int c = 0; c < width; c++)
                {

                    bool matches = true;
                    for (int i = 0; i < ranges.Length; i++)
                    {
                        AreaEval aeRange = ranges[i];
                        IMatchPredicate mp = predicates[i];

                        if (!mp.Matches(aeRange.GetRelativeValue(r, c)))
                        {
                            matches = false;
                            break;
                        }

                    }

                    if (matches)
                    { // sum only if all of the corresponding criteria specified are true for that cell.
                        var result = Accumulate(aeAvg, r, c);
                        if(result == null) continue;
                        sum += result.Value;
                        valuesCount++;
                    }
                }
            }

            if (valuesCount <= 0)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            return sum / valuesCount;
        }

        private static double? Accumulate(AreaEval aeSum, int relRowIndex,
                int relColIndex)
        {

            ValueEval addend = aeSum.GetRelativeValue(relRowIndex, relColIndex);
            if (addend is NumberEval)
            {
                return ((NumberEval)addend).NumberValue;
            }
            // everything else (including string and boolean values) dont count
            return null;
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
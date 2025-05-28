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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    /// <summary>
    /// Base class for SUMIFS() and COUNTIFS() functions, as they share much of the same logic, 
    /// the difference being the source of the totals.
    /// </summary>
    public abstract class Baseifs : FreeRefFunction
    {
        /// <summary>
        /// Implementations must be stateless.
        /// return true if there should be a range argument before the criteria pairs
        /// </summary>
        protected abstract bool HasInitialRange { get; }

        protected interface IAggregator
        {
            void AddValue(ValueEval value);
            ValueEval GetResult();
        }

        protected abstract IAggregator CreateAggregator();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            bool hasInitialRange = HasInitialRange;
            int firstCriteria = hasInitialRange ? 1 : 0;

            if(args.Length < (2 + firstCriteria) || args.Length % 2 != firstCriteria)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                AreaEval sumRange = null;
                if(hasInitialRange)
                {
                    sumRange = ConvertRangeArg(args[0]);
                }

                // collect pairs of ranges and criteria
                AreaEval[] ae = new AreaEval[(args.Length - firstCriteria) / 2];
                IMatchPredicate[] mp = new IMatchPredicate[ae.Length];
                for(int i = firstCriteria, k = 0; i < (args.Length - 1); i += 2, k++)
                {
                    ae[k] = ConvertRangeArg(args[i]);

                    mp[k] = Countif.CreateCriteriaPredicate(args[i + 1], ec.RowIndex, ec.ColumnIndex);
                }

                ValidateCriteriaRanges(sumRange, ae);
                ValidateCriteria(mp);

                return AggregateMatchingCells(CreateAggregator(), sumRange, ae, mp);
            }
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        /// <summary>
        /// Verify that each <c>criteriaRanges</c> argument contains the same number of rows and columns
        /// including the <c>sumRange</c> argument if present
        /// </summary>
        /// <param name="sumRange">if used, it must match the shape of the criteriaRanges</param>
        /// <param name="criteriaRanges">criteriaRanges to check</param>
        /// <exception cref="EvaluationException">throws EvaluationException if the ranges do not match.</exception>
        protected internal static void ValidateCriteriaRanges(AreaEval sumRange, AreaEval[] criteriaRanges)
        {
            int h = criteriaRanges[0].Height;
            int w = criteriaRanges[0].Width;

            if(sumRange != null
                    && (sumRange.Height != h
                    || sumRange.Width != w))
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            }

            foreach(AreaEval r in criteriaRanges)
            {
                if(r.Height != h ||
                        r.Width != w)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
            }
        }

        /// <summary>
        /// Verify that each <c>criteria</c> predicate is valid, i.e. not an error
        /// </summary>
        /// <param name="criteria">criteria to check</param>
        /// <exception cref="EvaluationException">throws EvaluationException if there are criteria which resulted in Errors.</exception>
        protected internal static void ValidateCriteria(IMatchPredicate[] criteria)
        {
            foreach(IMatchPredicate predicate in criteria)
            {
                // check for errors in predicate and return immediately using this error code
                if(predicate is Countif.ErrorMatcher matcher)
                {
                    throw new EvaluationException(
                        ErrorEval.ValueOf(matcher.Value));
                }
            }
        }


        /**
         * @param sumRange  the range to sum, if used (uses 1 for each match if not present)
         * @param ranges  criteria ranges
         * @param predicates  array of predicates, a predicate for each value in <code>ranges</code>
         * @return the computed value
         * @throws EvaluationException if there is an issue with eval
         */
        protected static ValueEval AggregateMatchingCells(IAggregator aggregator, AreaEval sumRange, AreaEval[] ranges, IMatchPredicate[] predicates)
        {
            int height = ranges[0].Height;
            int width = ranges[0].Width;

            for(int r = 0; r < height; r++)
            {
                for(int c = 0; c < width; c++)
                {
                    bool matches = true;
                    for(int i = 0; i < ranges.Length; i++)
                    {
                        AreaEval aeRange = ranges[i];
                        IMatchPredicate mp = predicates[i];

                        if(mp == null || !mp.Matches(aeRange.GetRelativeValue(r, c)))
                        {
                            matches = false;
                            break;
                        }
                    }

                    if(matches)
                    { // aggregate only if all of the corresponding criteria specified are true for that cell.
                        if(sumRange != null)
                        {
                            ValueEval value = sumRange.GetRelativeValue(r, c);
                            if(value is ErrorEval eval)
                            {
                                throw new EvaluationException(eval);
                            }
                            aggregator.AddValue(value);
                        }
                        else
                        {
                            aggregator.AddValue(null);
                        }
                    }
                }
            }

            return aggregator.GetResult();
        }


        protected internal static AreaEval ConvertRangeArg(ValueEval eval)
        {
            if(eval is AreaEval areaEval)
            {
                return areaEval;
            }
            if(eval is RefEval refEval)
            {
                return refEval.Offset(0, 0, 0, 0);
            }
            throw new EvaluationException(ErrorEval.VALUE_INVALID);
        }
    }
}

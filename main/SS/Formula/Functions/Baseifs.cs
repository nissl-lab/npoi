using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public abstract class Baseifs : FreeRefFunction
    {
        public abstract bool HasInitialRange();

        public interface IAggregator
        {
            void AddValue(ValueEval d);
            ValueEval GetResult();
        }

        public abstract IAggregator CreateAggregator();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            bool hasInitialRange = HasInitialRange();
            int firstCriteria = hasInitialRange ? 1 : 0;

            if (args.Length < (2 + firstCriteria) || args.Length % 2 != firstCriteria)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                AreaEval sumRange = null;
                if (hasInitialRange)
                {
                    sumRange = convertRangeArg(args[0]);
                }

                // collect pairs of ranges and criteria
                AreaEval[] ae = new AreaEval[(args.Length - firstCriteria) / 2];
                IMatchPredicate[] mp = new IMatchPredicate[ae.Length];
                for (int i = firstCriteria, k = 0; i < (args.Length - 1); i += 2, k++)
                {
                    ae[k] = convertRangeArg(args[i]);

                    mp[k] = Countif.CreateCriteriaPredicate(args[i + 1], ec.RowIndex, ec.ColumnIndex);                    
                }

                validateCriteriaRanges(sumRange, ae);
                validateCriteria(mp);

                return aggregateMatchingCells(CreateAggregator(), sumRange, ae, mp);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        /**
         * Verify that each <code>criteriaRanges</code> argument contains the same number of rows and columns
         * including the <code>sumRange</code> argument if present
         * @param sumRange if used, it must match the shape of the criteriaRanges
         * @param criteriaRanges to check
         * @throws EvaluationException if the ranges do not match.
         */
        private static void validateCriteriaRanges(AreaEval sumRange, AreaEval[] criteriaRanges)
        {
            int h = criteriaRanges[0].Height;
            int w = criteriaRanges[0].Width;

            if (sumRange != null
                    && (sumRange.Height != h
                    || sumRange.Width != w))
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            }

            foreach (AreaEval r in criteriaRanges)
            {
                if (r.Height != h ||
                        r.Width != w)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
            }
        }

        /**
        * Verify that each <code>criteria</code> predicate is valid, i.e. not an error
        * @param criteria to check
        *
        * @throws EvaluationException if there are criteria which resulted in Errors.
        */
        private static void validateCriteria(IMatchPredicate[] criteria)
        {
            foreach (IMatchPredicate predicate in criteria)
            {
                // check for errors in predicate and return immediately using this error code
                if (predicate is Countif.ErrorMatcher)
                {
                    throw new EvaluationException(
                        ErrorEval.ValueOf(((NPOI.SS.Formula.Functions.Countif.ErrorMatcher)predicate).Value));
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
        private static ValueEval aggregateMatchingCells(IAggregator aggregator, AreaEval sumRange, AreaEval[] ranges, IMatchPredicate[] predicates)
        {
            int height = ranges[0].Height;
            int width = ranges[0].Width;

            for (int r = 0; r < height; r++) 
            {
                for (int c = 0; c < width; c++) 
                {
                    bool matches = true;
                    for (int i = 0; i < ranges.Length; i++) 
                    {
                        AreaEval aeRange = ranges[i];
                        IMatchPredicate mp = predicates[i];

                        if (mp == null || !mp.Matches(aeRange.GetRelativeValue(r, c))) 
                        {
                            matches = false;
                            break;
                        }
                    }

                    if (matches)
                    { // aggregate only if all of the corresponding criteria specified are true for that cell.
                        if (sumRange != null)
                        {
                            ValueEval value = sumRange.GetRelativeValue(r, c);
                            if (value is ErrorEval) 
                            {
                                throw new EvaluationException((ErrorEval)value);
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


        protected static AreaEval convertRangeArg(ValueEval eval)
        {
            if (eval is AreaEval) {
                return (AreaEval) eval;
            }
            if (eval is RefEval) {
                return ((RefEval) eval).Offset(0, 0, 0, 0);
            }
            throw new EvaluationException(ErrorEval.VALUE_INVALID);
        }
    }
}

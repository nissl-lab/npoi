using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Atp
{
    public class XMatchFunction : FreeRefFunction
    {

        public static FreeRefFunction instance = new XMatchFunction(ArgumentsEvaluator.instance);

        private ArgumentsEvaluator evaluator;

        private XMatchFunction(ArgumentsEvaluator anEvaluator)
        {
            // enforces singleton
            this.evaluator = anEvaluator;
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            int srcRowIndex = ec.RowIndex;
            int srcColumnIndex = ec.ColumnIndex;
            return _evaluate(args, srcRowIndex, srcColumnIndex);
        }

        private ValueEval _evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length < 2)
            {
                return ErrorEval.VALUE_INVALID;
            }
            LookupUtils.MatchMode matchMode = LookupUtils.MatchMode.ExactMatch;
            if (args.Length > 2)
            {
                try
                {
                    ValueEval matchModeValue = OperandResolver.GetSingleValue(args[2], srcRowIndex, srcColumnIndex);
                    int matchInt = OperandResolver.CoerceValueToInt(matchModeValue);
                    matchMode = LookupUtils.GetMatchMode(matchInt);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
                catch
                {
                    return ErrorEval.VALUE_INVALID;
                }
            }
            LookupUtils.SearchMode searchMode = LookupUtils.SearchMode.IterateForward;
            if (args.Length > 3)
            {
                try
                {
                    ValueEval searchModeValue = OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex);
                    int searchInt = OperandResolver.CoerceValueToInt(searchModeValue);
                    searchMode = LookupUtils.GetSearchMode(searchInt);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
                catch
                {
                    return ErrorEval.VALUE_INVALID;
                }
            }
            return evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], matchMode, searchMode);
        }

        private ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval lookupEval, ValueEval indexEval,
                                   LookupUtils.MatchMode matchMode, LookupUtils.SearchMode searchMode)
        {
            try
            {
                ValueEval lookupValue = OperandResolver.GetSingleValue(lookupEval, srcRowIndex, srcColumnIndex);
                TwoDEval tableArray = LookupUtils.ResolveTableArrayArg(indexEval);
                ValueVector vector;
                if (tableArray.IsColumn)
                {
                    vector = LookupUtils.CreateColumnVector(tableArray, 0);
                }
                else
                {
                    vector = LookupUtils.CreateRowVector(tableArray, 0);
                }
                int matchedIdx = LookupUtils.XlookupIndexOfValue(lookupValue, vector, matchMode, searchMode);
                return new NumberEval((double)matchedIdx + 1);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }
}


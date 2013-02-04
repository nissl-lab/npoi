using System;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    public class SearchFind : Var2or3ArgFunction
    {

        private bool _isCaseSensitive;

        public SearchFind(bool isCaseSensitive)
        {
            _isCaseSensitive = isCaseSensitive;
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            try
            {
                String needle = TextFunction.EvaluateStringArg(arg0, srcRowIndex, srcColumnIndex);
                String haystack = TextFunction.EvaluateStringArg(arg1, srcRowIndex, srcColumnIndex);
                return Eval(haystack, needle, 0);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2)
        {
            try
            {
                String needle = TextFunction.EvaluateStringArg(arg0, srcRowIndex, srcColumnIndex);
                String haystack = TextFunction.EvaluateStringArg(arg1, srcRowIndex, srcColumnIndex);
                // evaluate third arg and convert from 1-based to 0-based index
                int startpos = TextFunction.EvaluateIntArg(arg2, srcRowIndex, srcColumnIndex) - 1;
                if (startpos < 0)
                {
                    return ErrorEval.VALUE_INVALID;
                }
                return Eval(haystack, needle, startpos);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        private ValueEval Eval(String haystack, String needle, int startIndex)
        {
            int result;
            if (_isCaseSensitive)
            {
                result = haystack.IndexOf(needle, startIndex, StringComparison.CurrentCulture);
            }
            else
            {
                //result = haystack.ToUpper().IndexOf(needle.ToUpper(), startIndex);
                result = haystack.IndexOf(needle, startIndex, StringComparison.CurrentCultureIgnoreCase);
            }
            if (result == -1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return new NumberEval(result + 1);
        }
    }
}

/*
2020-03-20 Buzz Weetman
Custom created because NPOI doesn't implement CELL.
I am specifically fixing a case where it is used as CELL("col",W2)
The COLUMN function will do that same thing as COLUMN(W2)
But to avoid having to change existing templates, this implementation has been made.
It only handles the "col" and "row" parameters.
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;

    public class Cell : Function2Arg
    {
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            string s0 = TextFunction.EvaluateStringArg(arg0, srcRowIndex, srcColumnIndex);

            // Only "col" and "row" are implemented

            int rnum;

            if(string.Equals(s0, "col", System.StringComparison.InvariantCultureIgnoreCase))
            {
                // "col"
                if(arg1 is AreaEval)
                {
                    rnum = ((AreaEval) arg1).FirstColumn;
                }
                else if(arg1 is RefEval)
                {
                    rnum = ((RefEval) arg1).Column;
                }
                else
                {
                    // anything else is not valid argument
                    return ErrorEval.VALUE_INVALID;
                }
            }
            else if(string.Equals(s0, "row", System.StringComparison.InvariantCultureIgnoreCase))
            {
                // "row"
                if(arg1 is AreaEval)
                {
                    rnum = ((AreaEval) arg1).FirstRow;
                }
                else if(arg1 is RefEval)
                {
                    rnum = ((RefEval) arg1).Row;
                }
                else
                {
                    // anything else is not valid argument
                    return ErrorEval.VALUE_INVALID;
                }
            }
            else
            {
                //https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf
                // anything else is not valid argument, until we implement more
                return ErrorEval.VALUE_INVALID;
            }

            return new NumberEval(rnum + 1);
        }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length == 2)
            {
                return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1]);
            }

            return ErrorEval.VALUE_INVALID;
        }
    }
}
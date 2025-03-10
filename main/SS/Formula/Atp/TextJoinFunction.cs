using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Atp
{
    public class TextJoinFunction : FreeRefFunction
    {
        public static FreeRefFunction instance = new TextJoinFunction(ArgumentsEvaluator.instance);

        private ArgumentsEvaluator evaluator;
        private TextJoinFunction(ArgumentsEvaluator anEvaluator)
        {
            // enforces singleton
            this.evaluator = anEvaluator;
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            /*
                     * Must be at least three arguments:
                     *  - delimiter    Delimiter for joining text arguments
                     *  - ignoreEmpty  If true, empty strings will be ignored in the join
                     *  - text1        First value to be evaluated as text and joined
                     *  - text2, etc.  Optional additional values to be evaluated and joined
                     */

            // Make sure we have at least one text value, and at most 252 text values, as documented at:
            // https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c?ui=en-us&rs=en-us&ad=us
            if (args.Length < 3 || args.Length > 254)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int srcRowIndex = ec.RowIndex;
            int srcColumnIndex = ec.ColumnIndex;

            try
            {
                // Get the delimiter argument
                List<ValueEval> delimiterArgs = GetValues(args[0], srcRowIndex, srcColumnIndex, true);

                // Get the boolean ignoreEmpty argument
                ValueEval ignoreEmptyArg = OperandResolver.GetSingleValue(args[1], srcRowIndex, srcColumnIndex);
                bool ignoreEmpty = (bool)OperandResolver.CoerceValueToBoolean(ignoreEmptyArg, false);

                // Get a list of string values for each text argument
                List<string> textValues = new List<string>();

                for (int i = 2; i < args.Length; i++)
                {
                    List<ValueEval> textArgs = GetValues(args[i], srcRowIndex, srcColumnIndex, false);
                    foreach (ValueEval textArg in textArgs)
                    {
                        String textValue = OperandResolver.CoerceValueToString(textArg);

                        // If we're not ignoring empty values or if our value is not empty, add it to the list
                        if (!ignoreEmpty || (textValue != null && textValue.Length > 0))
                        {
                            textValues.Add(textValue);
                        }
                    }
                }

                // Join the list of values with the specified delimiter and return
                if (delimiterArgs.Count == 0)
                {
                    return new StringEval(String.Join("", textValues));
                }
                else if (delimiterArgs.Count == 1)
                {
                    String delimiter = LaxValueToString(delimiterArgs[0]);
                    return new StringEval(String.Join(delimiter, textValues));
                }
                else
                {
                    //https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c
                    //see example 3 to see why this is needed
                    List<string> delimiters = new List<string>();
                    foreach (ValueEval delimiterArg in delimiterArgs)
                    {
                        delimiters.Add(LaxValueToString(delimiterArg));
                    }
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < textValues.Count; i++)
                    {
                        if (i > 0)
                        {
                            int delimiterIndex = (i - 1) % delimiters.Count;
                            sb.Append(delimiters[delimiterIndex]);
                        }
                        sb.Append(textValues[i]);
                    }
                    return new StringEval(sb.ToString());
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private String LaxValueToString(ValueEval eval)
        {
            return (eval is MissingArgEval) ? "" : OperandResolver.CoerceValueToString(eval);
        }

        //https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c
        //in example 3, the delimiter is defined by a large area but only the last row of that area seems to be used
        //this is why lastRowOnly is supported
        private List<ValueEval> GetValues(ValueEval eval, int srcRowIndex, int srcColumnIndex, bool lastRowOnly)
        {
            if (eval is AreaEval ae)
            {
                List<ValueEval> list = new List<ValueEval>();
                int startRow = lastRowOnly ? ae.LastRow : ae.FirstRow;
                for (int r = startRow; r <= ae.LastRow; r++)
                {
                    for (int c = ae.FirstColumn; c <= ae.LastColumn; c++)
                    {
                        list.Add(OperandResolver.GetSingleValue(ae.GetAbsoluteValue(r, c), r, c));
                    }
                }
                return list;
            }
            else
            {
                return new List<ValueEval>() { OperandResolver.GetSingleValue(eval, srcRowIndex, srcColumnIndex) };
            }
        }
    }
}

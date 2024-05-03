using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;

namespace NPOI.SS.Formula.Functions
{
    /**
	 * An implementation of the TEXT function
	 * TEXT returns a number value formatted with the given number formatting string. 
	 * This function is not a complete implementation of the Excel function, but
	 *  handles most of the common cases. All work is passed down to 
	 *  {@link DataFormatter} to be done, as this works much the same as the
	 *  display focused work that that does. 
	 */
    public class Text : Fixed2ArgFunction
    {
        public static DataFormatter Formatter { get; set; } = new();

        /// <summary>
        /// An implementation of the TEXT function <br/>
        /// TEXT returns a number value formatted with the given number formatting string. <br/>
        /// This function is not a complete implementation of the Excel function, but <br/>
        /// handles most of the common cases. All work is passed down to <br/>
        /// <see cref="DataFormatter"/> to be done, as this works much the same as the <br/>
        /// display focused work that does.
        /// </summary>
        /// <param name="srcRowIndex"></param>
        /// <param name="srcColumnIndex"></param>
        /// <param name="arg0"></param>
        /// <param name="arg1"></param>
        /// <returns></returns>
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            ValueEval valueEval;

            try
            {
                ValueEval valueVe = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                ValueEval formatVe = OperandResolver.GetSingleValue(arg1, srcRowIndex, srcColumnIndex);

                try
                {
                    double valueDouble = double.NaN;
                    string evaluated = null;

                    if (valueVe == BlankEval.instance)
                    {
                        valueDouble = 0.0;
                    }
                    else if (valueVe is BoolEval boolEval) 
                    {
                        evaluated = boolEval.StringValue;
                    } 
                    else if (valueVe is NumericValueEval numericEval) 
                    {
                        valueDouble = numericEval.NumberValue;
                                                 
                    } 
                    else if (valueVe is StringEval stringEval) 
                    {
                        evaluated = stringEval.StringValue;
                        valueDouble = OperandResolver.ParseDouble(evaluated);
                    }

                    if(!double.IsNaN(valueDouble))
                    {
                        string format = FormatPatternValueEval2String(formatVe);
                        evaluated = Formatter.FormatRawCellContents(valueDouble, -1, format);
                    }

                    valueEval = new StringEval(evaluated);
                }
                catch(Exception)
                {
                    valueEval = ErrorEval.VALUE_INVALID;
                }
            }
            catch(EvaluationException e)
            {
                valueEval = e.GetErrorEval();
            }

            return valueEval;
        }

        /// <summary>
        /// Using it instead of <see cref="OperandResolver.CoerceValueToString(ValueEval)"/> in order to handle booleans differently.
        /// </summary>
        /// <param name="ve"></param>
        /// <returns>Pattern value eval formatted to string</returns>
        /// <exception cref="ArgumentException"></exception>
        private string FormatPatternValueEval2String(ValueEval ve)
        {
            string format;

            if (ve is not BoolEval && ve is StringValueEval sve) 
            {
                format = sve.StringValue;
            } 
            else if (ve == BlankEval.instance)
            {
                format = "";
            }
            else
            {
                throw new ArgumentException("Unexpected eval class (" + ve.GetType().Name + ")");
            }
            
            return format;
        }
    }
}

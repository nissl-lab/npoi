using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Implementation for Excel REPT () function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>REPT  </b>(<b>text</b>,<b>number_times</b> )<br/>
     * <p/>
     * Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string.
     *
     * text : text The text that you want to repeat.
     * number_times:	A positive number specifying the number of times to repeat text.
     *
     * If number_times is 0 (zero), REPT returns "" (empty text).
     * If this argument contains a decimal value, this function ignores the numbers to the right side of the decimal point.
     *
     * The result of the REPT function cannot be longer than 32,767 characters, or REPT returns #VALUE!.
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Rept : Fixed2ArgFunction
    {


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval text, ValueEval number_times)
        {

            ValueEval veText1;
            try
            {
                veText1 = OperandResolver.GetSingleValue(text, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            String strText1 = OperandResolver.CoerceValueToString(veText1);
            double numberOfTime = 0;
            try
            {
                numberOfTime = OperandResolver.CoerceValueToDouble(number_times);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int numberOfTimeInt = (int)numberOfTime;
            StringBuilder strb = new StringBuilder(strText1.Length * numberOfTimeInt);
            for (int i = 0; i < numberOfTimeInt; i++)
            {
                strb.Append(strText1);
            }

            if (strb.ToString().Length > 32767)
            {
                return ErrorEval.VALUE_INVALID;
            }

            return new StringEval(strb.ToString());
        }
    }
}

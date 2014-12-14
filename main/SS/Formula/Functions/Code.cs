using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Implementation for Excel CODE () function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>CODE   </b>(<b>text</b> )<br/>
     * <p/>
     * Returns a numeric code for the first character in a text string. The returned code corresponds to the character set used by your computer.
     * <p/>
     * text The text for which you want the code of the first character.
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Code : Fixed1ArgFunction
    {

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval textArg)
        {

            ValueEval veText1;
            try
            {
                veText1 = OperandResolver.GetSingleValue(textArg, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            String text = OperandResolver.CoerceValueToString(veText1);

            if (text.Length == 0)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int code = (int)text[0];

            return new StringEval(code.ToString());
        }
    }
}

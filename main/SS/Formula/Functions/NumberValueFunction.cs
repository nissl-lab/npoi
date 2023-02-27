using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace NPOI.SS.Formula.Functions
{
    public class NumberValueFunction : FreeRefFunction
    {
        public static FreeRefFunction instance = new NumberValueFunction();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            string decSep = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            string groupSep = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator;
            String text = null;
            Double result = Double.NaN;
            ValueEval v1 = null;
            ValueEval v2 = null;
            ValueEval v3 = null;
            try
            {
                if (args.Length == 1)
                {
                    v1 = OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex);
                    text = OperandResolver.CoerceValueToString(v1);
                }
                else if (args.Length == 2)
                {
                    v1 = OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex);
                    v2 = OperandResolver.GetSingleValue(args[1], ec.RowIndex, ec.ColumnIndex);
                    text = OperandResolver.CoerceValueToString(v1);
                    decSep = OperandResolver.CoerceValueToString(v2).Substring(0, 1); //If multiple characters are used in the Decimal_separator or Group_separator arguments, only the first character is used.
                }
                else if (args.Length == 3)
                {
                    v1 = OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex);
                    v2 = OperandResolver.GetSingleValue(args[1], ec.RowIndex, ec.ColumnIndex);
                    v3 = OperandResolver.GetSingleValue(args[2], ec.RowIndex, ec.ColumnIndex);
                    text = OperandResolver.CoerceValueToString(v1);
                    decSep = OperandResolver.CoerceValueToString(v2).Substring(0, 1); //If multiple characters are used in the Decimal_separator or Group_separator arguments, only the first character is used.
                    groupSep = OperandResolver.CoerceValueToString(v3).Substring(0, 1); //If multiple characters are used in the Decimal_separator or Group_separator arguments, only the first character is used.
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            if (text == "")
                text = "0"; //If an empty string ("") is specified as the Text argument, the result is 0.
            text = text.Replace(" ", ""); //Empty spaces in the Text argument are ignored, even in the middle of the argument. For example, " 3 000 " is returned as 3000.
            String[] parts = text.Split(new string[] { decSep }, StringSplitOptions.RemoveEmptyEntries);
            String sigPart = "";
            String decPart = "";
            if (parts.Length > 2) return ErrorEval.VALUE_INVALID; //If a decimal separator is used more than once in the Text argument, NUMBERVALUE returns the #VALUE! error value.
            if (parts.Length > 1)
            {
                sigPart = parts[0];
                decPart = parts[1];
                if (decPart.Contains(groupSep)) return ErrorEval.VALUE_INVALID; //If the group separator occurs after the decimal separator in the Text argument, NUMBERVALUE returns the #VALUE! error value.
                sigPart = sigPart.Replace(groupSep, ""); //If the group separator occurs before the decimal separator in the Text argument , the group separator is ignored.
                text = sigPart + "." + decPart;
            }
            else if (parts.Length > 0)
            {
                sigPart = parts[0];
                sigPart = sigPart.Replace(groupSep, ""); //If the group separator occurs before the decimal separator in the Text argument , the group separator is ignored.
                text = sigPart;
            }
            // If the Text argument ends in one or more percent signs(%), they are used in the calculation of the result.
            //Multiple percent signs are additive if they are used in the Text argument just as they are if they are used in a formula.
            //For example, =NUMBERVALUE("9%%") returns the same result (0.0009) as the formula =9%%.
            int countPercent = 0;
            while (text.EndsWith("%"))
            {
                countPercent++;
                text = text.Substring(0, text.Length - 1);
            }

            try
            {
                result = Double.Parse(text);
                result = result / Math.Pow(100, countPercent); //If the Text argument ends in one or more percent signs (%), they are used in the calculation of the result.
                checkValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            catch
            {
                return ErrorEval.VALUE_INVALID; //If any of the arguments are not valid, NUMBERVALUE returns the #VALUE! error value.
            }

            return new NumberEval(result);
        }
        private static void checkValue(double result)
        {
            if (Double.IsNaN(result) || Double.IsInfinity(result)) {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }
    }
}

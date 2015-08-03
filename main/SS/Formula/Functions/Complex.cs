using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Implementation for Excel COMPLEX () function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>COMPLEX   </b>(<b>real_num</b>,<b>i_num</b>,<b>suffix </b> )<br/>
     * <p/>
     * Converts real and imaginary coefficients into a complex number of the form x + yi or x + yj.
     * <p/>
     * <p/>
     * All complex number functions accept "i" and "j" for suffix, but neither "I" nor "J".
     * Using uppercase results in the #VALUE! error value. All functions that accept two
     * or more complex numbers require that all suffixes match.
     * <p/>
     * <b>real_num</b> The real coefficient of the complex number.
     * If this argument is nonnumeric, this function returns the #VALUE! error value.
     * <p/>
     * <p/>
     * <b>i_num</b> The imaginary coefficient of the complex number.
     * If this argument is nonnumeric, this function returns the #VALUE! error value.
     * <p/>
     * <p/>
     * <b>suffix</b> The suffix for the imaginary component of the complex number.
     * <ul>
     * <li>If omitted, suffix is assumed to be "i".</li>
     * <li>If suffix is neither "i" nor "j", COMPLEX returns the #VALUE! error value.</li>
     * </ul>
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Complex : Var2or3ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction Instance = new Complex();

        public static String DEFAULT_SUFFIX = "i";
        public static String SUPPORTED_SUFFIX = "j";

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval real_num, ValueEval i_num)
        {
            return this.Evaluate(srcRowIndex, srcColumnIndex, real_num, i_num, new StringEval(DEFAULT_SUFFIX));
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval real_num, ValueEval i_num, ValueEval suffix)
        {
            ValueEval veText1;
            try
            {
                veText1 = OperandResolver.GetSingleValue(real_num, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            double realNum = 0;
            try
            {
                realNum = OperandResolver.CoerceValueToDouble(veText1);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }

            ValueEval veINum;
            try
            {
                veINum = OperandResolver.GetSingleValue(i_num, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            double realINum = 0;
            try
            {
                realINum = OperandResolver.CoerceValueToDouble(veINum);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }

            String suffixValue = OperandResolver.CoerceValueToString(suffix);
            if (suffixValue.Length == 0)
            {
                suffixValue = DEFAULT_SUFFIX;
            }
            if (suffixValue.Equals(DEFAULT_SUFFIX.ToUpper()) || suffixValue.Equals(SUPPORTED_SUFFIX.ToUpper()))
            {
                return ErrorEval.VALUE_INVALID;
            }
            if (!(suffixValue.Equals(DEFAULT_SUFFIX) || suffixValue.Equals(SUPPORTED_SUFFIX)))
            {
                return ErrorEval.VALUE_INVALID;
            }

            StringBuilder strb = new StringBuilder("");
            if (realNum != 0)
            {
                if (isDoubleAnInt(realNum))
                {
                    strb.Append((int)realNum);
                }
                else
                {
                    strb.Append(realNum);
                }
            }
            if (realINum != 0)
            {
                if (strb.Length != 0)
                {
                    if (realINum > 0)
                    {
                        strb.Append("+");
                    }
                }

                if (realINum != 1 && realINum != -1)
                {
                    if (isDoubleAnInt(realINum))
                    {
                        strb.Append((int)realINum);
                    }
                    else
                    {
                        strb.Append(realINum);
                    }
                }

                strb.Append(suffixValue);
            }

            return new StringEval(strb.ToString());
        }

        private bool isDoubleAnInt(double number)
        {
            return (number == Math.Floor(number)) && !Double.IsInfinity(number);
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 2)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
            }
            if (args.Length == 3)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1], args[2]);
            }

            return ErrorEval.VALUE_INVALID;
        }
    }
}

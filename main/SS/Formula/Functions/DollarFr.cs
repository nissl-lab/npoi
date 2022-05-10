using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class DollarFr : Fixed2ArgFunction, FreeRefFunction
    {
        public static FreeRefFunction instance = new DollarFr();
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1, ValueEval arg2)
        {
            try
            {
                Double number1 = evaluateValue(arg1, srcRowIndex, srcColumnIndex);
                if (Double.IsNaN(number1))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                Double number2 = evaluateValue(arg2, srcRowIndex, srcColumnIndex);
                if (Double.IsNaN(number2))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                int fraction = (int)number2;
                if (fraction < 0)
                {
                    return ErrorEval.NUM_ERROR;
                }
                else if (fraction == 0)
                {
                    return ErrorEval.DIV_ZERO;
                }
                int fractionLength = fraction.ToString().Length;

                bool negative = false;
                long valueLong = (long)number1;
                if (valueLong < 0)
                {
                    negative = true;
                    valueLong = -valueLong;
                    number1 = -number1;
                }

                double valueFractional = number1 - valueLong;
                if (valueFractional == 0.0)
                {
                    return new NumberEval(valueLong);
                }

                double result = valueFractional * fraction / Math.Pow(10, fractionLength)+valueLong;

                if (negative)
                {
                    result = result*-1;
                }

                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 2)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
            }

            return ErrorEval.VALUE_INVALID;
        }
        private static Double evaluateValue(ValueEval arg, int srcRowIndex, int srcColumnIndex)
        {
            ValueEval veText = OperandResolver.GetSingleValue(arg, srcRowIndex, srcColumnIndex);
            String strText1 = OperandResolver.CoerceValueToString(veText);
            return OperandResolver.ParseDouble(strText1);
        }
}
}

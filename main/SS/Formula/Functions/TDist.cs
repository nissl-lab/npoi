using MathNet.Numerics.Distributions;
using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class TDist : Fixed3ArgFunction, FreeRefFunction
    {
        public static TDist instance = new TDist();
        static double tdistOneTail(double x, int degreesOfFreedom)
        {
            return 1 - StudentT.CDF(0, 1, degreesOfFreedom, x);
        }

        static double tdistTwoTails(double x, int degreesOfFreedom)
        {
            return 2 * tdistOneTail(x, degreesOfFreedom);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1, ValueEval arg2, ValueEval arg3)
        {
            try
            {
                Double number1 = evaluateValue(arg1, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(number1))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                else if (number1 < 0)
                {
                    return ErrorEval.NUM_ERROR;
                }
                Double number2 = evaluateValue(arg2, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(number2))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                int degreesOfFreedom = (int)number2;
                if (degreesOfFreedom < 1)
                {
                    return ErrorEval.NUM_ERROR;
                }
                Double number3 = evaluateValue(arg3, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(number3))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                int tails = (int)number3;
                if (!(tails == 1 || tails == 2))
                {
                    return ErrorEval.NUM_ERROR;
                }

                if (tails == 2)
                {
                    return new NumberEval(tdistTwoTails(number1, degreesOfFreedom));
                }

                return new NumberEval(tdistOneTail(number1, degreesOfFreedom));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 3)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1], args[2]);
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

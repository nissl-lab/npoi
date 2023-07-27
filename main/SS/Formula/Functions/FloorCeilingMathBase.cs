using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public abstract class FloorCeilingMathBase : FreeRefFunction
    {
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
            => args.Length switch
            {
                1 => Evaluate(0, 0, args[0]),
                2 => Evaluate(0, 0, args[0], args[1]),
                3 => Evaluate(0, 0, args[0], args[1], args[2]),
                _ => ErrorEval.VALUE_INVALID
            };
        private ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            return Evaluate(srcRowIndex, srcColumnIndex, arg0, null, null);
        }
        private ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            return Evaluate(srcRowIndex, srcColumnIndex, arg0, arg1, null);
        }

        private ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1, ValueEval arg2)
        {
            try
            {
                double number = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double significance = arg1 is null ? 1.0 :
                    NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);

                bool? method = null;

                if (arg2 is not null)
                {
                    ValueEval ve = OperandResolver.GetSingleValue(arg2, srcRowIndex, srcColumnIndex);
                    method = OperandResolver.CoerceValueToBoolean(ve, false);
                }

                var result = Evaluate(number, significance, method ?? false);

                return result == 0.0 ? NumberEval.ZERO : new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        public double Evaluate(double number, double significance, bool mode)
        {
            if (significance == 0.0 || number == 0.0)
            {
                // FLOOR|CEILING.MATH 's behavior is different from FLOOR|CEILING
                // when significance is zero & number isn't 0, the MATH one returns 0 instead of #DIV/0.

                return 0.0;
            }

            if (number > 0 && significance < 0 || number < 0 && significance > 0)
            {
                // This is how Excel behaves
                significance = -significance;
            }

            return EvaluateMath(number, significance, mode);
        }

        protected abstract double EvaluateMajorDirection(double number);
        protected abstract double EvaluateAlternativeDirection(double number);
        private double EvaluateMath(double number, double significance, bool mode)
        {
            if (number >= 0)
            {
                // number is positive

                return EvaluateMajorDirection(number / significance) * significance;
            }
            else
            {
                // number is negative

                if (mode)
                {
                    // Towards zero for FLOOR && Away from zero for CEILING

                    return EvaluateAlternativeDirection(number / -significance) * -significance;
                }
                else
                {
                    // vice versa

                    return EvaluateMajorDirection(number / -significance) * -significance;
                }
            }
        }
    }
}

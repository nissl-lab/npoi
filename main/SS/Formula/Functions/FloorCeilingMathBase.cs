using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public abstract class FloorCeilingMathBase : FreeRefFunction
    {
        // Excel has an internal precision of 15 significant digits
        private const int SignificantThreshold = 15;

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
            => args.Length switch
            {
                1 => Evaluate(ec.RowIndex, ec.ColumnIndex, args[0]),
                2 => Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]),
                3 => Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1], args[2]),
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
                var number = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                var significance = arg1 is null ? 1.0 : NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);

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

            // Workaround without BigDecimal
            var numberToTest = number / significance;
            if (DoublePrecisionHelper.IsIntegerWithDigitsDropped(numberToTest, SignificantThreshold))
                return number;

            if (number > 0)
            {
                // mode is meaningless when number is positive
                return EvaluateMajorDirection(numberToTest) * significance;
            }
            else
            {
                if (mode)
                {
                    // Towards zero for FLOOR && Away from zero for CEILING
                    return EvaluateAlternativeDirection(-numberToTest) * -significance;
                }
                else
                {
                    // Vice versa
                    return EvaluateMajorDirection(-numberToTest) * -significance;
                }
            }
        }

        protected abstract double EvaluateMajorDirection(double number);
        protected abstract double EvaluateAlternativeDirection(double number);
    }
}

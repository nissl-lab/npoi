using MathNet.Numerics.Distributions;
using NPOI.SS.Formula.Eval;
using System;

namespace NPOI.SS.Formula.Functions
{
    public class TInv : Fixed2ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new TInv();

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1, ValueEval arg2)
        {
            var probability = EvaluateValue(arg1, srcRowIndex, srcColumnIndex);
            if(double.IsNaN(probability))
            {
                return ErrorEval.VALUE_INVALID;
            }
            var degreesOfFreedom = EvaluateValue(arg2, srcRowIndex, srcColumnIndex);
            if(double.IsNaN(degreesOfFreedom))
            {
                return ErrorEval.VALUE_INVALID;
            }

            StudentT studentT = new(0, 1, degreesOfFreedom);
            var result = studentT.InverseCumulativeDistribution(probability);
            return new NumberEval(result);
        }

        private static Double EvaluateValue(ValueEval arg, int srcRowIndex, int srcColumnIndex)
        {
            ValueEval veText = OperandResolver.GetSingleValue(arg, srcRowIndex, srcColumnIndex);
            String strText1 = OperandResolver.CoerceValueToString(veText);
            return OperandResolver.ParseDouble(strText1);
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if(args.Length == 2)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
            }
            return ErrorEval.VALUE_INVALID;
        }
    }
}

using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class NormInv:Fixed3ArgFunction,FreeRefFunction
    {
        public static NormInv instance = new NormInv();
        internal static double inverse(double probability, double mean, double stdev)
        {
            var normalDistribution = new MathNet.Numerics.Distributions.Normal(mean, stdev);
            return normalDistribution.InverseCumulativeDistribution(probability);
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1, ValueEval arg2, ValueEval arg3)
        {
            try
            {
                Double probability = evaluateValue(arg1, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(probability))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                else if (probability <= 0 || probability >= 1)
                {
                    return ErrorEval.NUM_ERROR;
                }
                Double mean = evaluateValue(arg2, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(mean))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                Double stdev = evaluateValue(arg3, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(stdev))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                else if (stdev <= 0)
                {
                    return ErrorEval.NUM_ERROR;
                }

                return new NumberEval(inverse(
                        probability, mean, stdev));
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

using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class NormDist : Fixed4ArgFunction, FreeRefFunction
    {
        public static NormDist instance = new NormDist();
        internal static double probability(double x, double mean, double stdev, bool cumulative)
        {
            var normalDistribution = new MathNet.Numerics.Distributions.Normal(mean, stdev);
            return cumulative ? normalDistribution.CumulativeDistribution(x) : normalDistribution.Density(x);
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1, ValueEval arg2, ValueEval arg3, ValueEval arg4)
        {
            try
            {
                Double xval = evaluateValue(arg1, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(xval))
                {
                    return ErrorEval.VALUE_INVALID;
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
                var cumulative = OperandResolver.CoerceValueToBoolean(arg4, false);
                if (cumulative==null)
                {
                    return ErrorEval.VALUE_INVALID;
                }

                return new NumberEval(probability(
                        xval, mean, stdev, (bool)cumulative));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 4)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1], args[2], args[3]);
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
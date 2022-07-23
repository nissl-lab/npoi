using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class BesselJ:Fixed2ArgFunction, FreeRefFunction
    {
        public static FreeRefFunction instance = new BesselJ();
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1, ValueEval arg2)
        {
            try
            {
                double xval = EvaluateValue(arg1, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(xval))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                Double orderDouble = EvaluateValue(arg2, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(orderDouble))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                int order = (int)orderDouble;
                if (order < 0)
                {
                    return ErrorEval.NUM_ERROR;
                }

                var result=MathNet.Numerics.SpecialFunctions.BesselJ(order, xval);
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
        private static Double EvaluateValue(ValueEval arg, int srcRowIndex, int srcColumnIndex)
        {
            ValueEval veText = OperandResolver.GetSingleValue(arg, srcRowIndex, srcColumnIndex);
            String strText1 = OperandResolver.CoerceValueToString(veText);
            return OperandResolver.ParseDouble(strText1);
        }
    }
}

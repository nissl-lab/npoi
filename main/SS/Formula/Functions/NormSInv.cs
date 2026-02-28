using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class NormSInv:Fixed1ArgFunction, FreeRefFunction
    {
        public static NormSInv instance = new NormSInv();
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1)
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

                return new NumberEval(NormInv.inverse(probability, 0, 1));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 1)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0]);
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

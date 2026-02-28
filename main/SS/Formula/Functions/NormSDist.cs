using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class NormSDist : Fixed1ArgFunction, FreeRefFunction
    {
        public static NormSDist instance = new NormSDist();
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg1)
        {
            try
            {
                Double xval = evaluateValue(arg1, srcRowIndex, srcColumnIndex);
                if (double.IsNaN(xval))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                else if (xval < 0)
                {
                    return ErrorEval.NUM_ERROR;
                }

                return new NumberEval(NormDist.probability(xval, 0, 1, true));
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

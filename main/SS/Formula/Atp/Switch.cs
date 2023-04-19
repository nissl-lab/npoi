using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Atp
{
    public class Switch : FreeRefFunction
    {
        public static FreeRefFunction instance = new Switch();

        private Switch()
        {
            // enforce singleton
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3) return ErrorEval.NA;

            ValueEval expression;
            try
            {
                expression = OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex);
            }
            catch
            {
                return ErrorEval.NA;
            }

            for (int i = 1; i < args.Length; i = i + 2)
            {

                try
                {
                    ValueEval value = OperandResolver.GetSingleValue(args[i], ec.RowIndex, ec.ColumnIndex);
                    ValueEval result = args[i + 1];
                    //ValueEval result = OperandResolver.getSingleValue(args[i+1],ec.RowIndex,ec.ColumnIndex) ;


                    ValueEval evaluate = (new EqualEval()).Evaluate(new ValueEval[] { expression, value }, ec.RowIndex, ec.ColumnIndex);
                    if (evaluate is BoolEval)
                    {
                        BoolEval boolEval = (BoolEval)evaluate;
                        bool booleanValue = boolEval.BooleanValue;
                        if (booleanValue)
                        {
                            return result;
                        }

                    }

                }
                catch (EvaluationException)
                {
                    return ErrorEval.NA;
                }

                if (i + 2 == args.Length - 1)
                {
                    //last value in args is the default one
                    return args[args.Length - 1];
                }

            }

            return ErrorEval.NA;
        }
    }
}
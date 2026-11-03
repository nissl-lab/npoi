using System;
using NPOI.SS.Formula.Functions;

namespace NPOI.SS.Formula.Eval
{

    public abstract class TwoOperandNumericOperation : Fixed2ArgFunction, IArrayFunction
    {
        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }
            ValueEval arg0 = args[0];
            ValueEval arg1 = args[1];

            int w1, h1, a1FirstCol = 0, a1FirstRow = 0;
            if (arg0 is AreaEval area0)
            {
                w1 = area0.Width;
                h1 = area0.Height;
                a1FirstCol = area0.FirstColumn;
                a1FirstRow = area0.FirstRow;
            }
            else if (arg0 is RefEval ref0)
            {
                w1 = 1; h1 = 1;
                a1FirstCol = ref0.Column;
                a1FirstRow = ref0.Row;
            }
            else
            {
                w1 = 1; h1 = 1;
            }

            int w2, h2, a2FirstCol = 0, a2FirstRow = 0;
            if (arg1 is AreaEval area1)
            {
                w2 = area1.Width;
                h2 = area1.Height;
                a2FirstCol = area1.FirstColumn;
                a2FirstRow = area1.FirstRow;
            }
            else if (arg1 is RefEval ref1)
            {
                w2 = 1; h2 = 1;
                a2FirstCol = ref1.Column;
                a2FirstRow = ref1.Row;
            }
            else
            {
                w2 = 1; h2 = 1;
            }

            int width = Math.Max(w1, w2);
            int height = Math.Max(h1, h2);

            ValueEval[] vals = new ValueEval[height * width];
            int idx = 0;
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    vals[idx++] = EvaluateOneArrayElement(
                        a1FirstRow + i, a1FirstCol + j,
                        a2FirstRow + i, a2FirstCol + j,
                        arg0, arg1);
                }
            }

            if (vals.Length == 1)
            {
                return vals[0];
            }

            return new CacheAreaEval(srcRowIndex, srcColumnIndex,
                    srcRowIndex + height - 1, srcColumnIndex + width - 1, vals);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            return EvaluateOneArrayElement(srcRowIndex, srcColumnIndex, srcRowIndex, srcColumnIndex, arg0, arg1);
        }

        // Evaluates a single element of an array operation independently, so that an error
        // produced by one element (e.g. an ErrorEval elsewhere in an array operand) cannot
        // poison the result of unrelated elements.
        private ValueEval EvaluateOneArrayElement(int row0, int col0, int row1, int col1, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                ValueEval ve0 = OperandResolver.GetSingleValue(arg0, row0, col0);
                ValueEval ve1 = OperandResolver.GetSingleValue(arg1, row1, col1);
                double d0 = OperandResolver.CoerceValueToDouble(ve0);
                double d1 = OperandResolver.CoerceValueToDouble(ve1);
                result = Evaluate(d0, d1);
                if (result == 0.0)
                { // this '==' matches +0.0 and -0.0
                    // Excel Converts -0.0 to +0.0 for '*', '/', '%', '+' and '^'
                    if (this is not SS.Formula.Eval.SubtractEval)
                    {
                        return NumberEval.ZERO;
                    }
                }
                if (Double.IsNaN(result) || Double.IsInfinity(result))
                {
                    return ErrorEval.NUM_ERROR;
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }

        public abstract double Evaluate(double d0, double d1);

        public static NPOI.SS.Formula.Functions.Function AddEval = new AddEval();
        public static NPOI.SS.Formula.Functions.Function DivideEval = new DivideEval();
        public static NPOI.SS.Formula.Functions.Function MultiplyEval = new MultiplyEval();
        public static NPOI.SS.Formula.Functions.Function PowerEval = new PowerEval();
        public static NPOI.SS.Formula.Functions.Function SubtractEval = new SubtractEval();
    }
}
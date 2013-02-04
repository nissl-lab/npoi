using System;
using NPOI.SS.Formula.Functions;

namespace NPOI.SS.Formula.Eval
{

    public abstract class TwoOperandNumericOperation : Fixed2ArgFunction
    {
        //public int Type
        //{
        //    get
        //    {
        //        // TODO - remove
        //        throw new Exception("obsolete code should not be called");
        //    }
        //}
        protected double SingleOperandEvaluate(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            return OperandResolver.CoerceValueToDouble(ve);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                double d0 = SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                result = Evaluate(d0, d1);
                if (result == 0.0)
                { // this '==' matches +0.0 and -0.0
                    // Excel Converts -0.0 to +0.0 for '*', '/', '%', '+' and '^'
                    if (!(this is SubtractEval))
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
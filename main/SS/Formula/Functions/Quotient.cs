using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.Eval;
namespace NPOI.SS.Formula.Functions
{
    /**
     * <p>Implementation for Excel QUOTIENT () function.</p>
     * <p>
     * <b>Syntax</b>:<br/> <b>QUOTIENT</b>(<b>Numerator</b>,<b>Denominator</b>)<br/>
     * </p>
     * <p>
     * Numerator     is the dividend.
     * Denominator     is the divisor.
     *
     * Returns the integer portion of a division. Use this function when you want to discard the remainder of a division.
     * </p>
     *
     * If either enumerator/denominator is non numeric, QUOTIENT returns the #VALUE! error value.
     * If denominator is Equals to zero, QUOTIENT returns the #DIV/0! error value.
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Quotient : Fixed2ArgFunction, FreeRefFunction
    {
        public static FreeRefFunction instance = new Quotient();
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval venumerator, ValueEval vedenominator)
        {

            double enumerator = 0;
            try
            {
                enumerator = OperandResolver.CoerceValueToDouble(venumerator);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }

            double denominator = 0;
            try
            {
                denominator = OperandResolver.CoerceValueToDouble(vedenominator);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }

            if (denominator == 0)
            {
                return ErrorEval.DIV_ZERO;
            }

            return new NumberEval(((int)(enumerator / denominator)));
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
        }
    }

}
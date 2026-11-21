using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public class Days : FreeRefFunction
    {
        public static Days Instance = new Days();

        private Days() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            return Days.Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
        }
        private static ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                var d0 = getDate(arg0, srcRowIndex, srcColumnIndex);
                var d1 = getDate(arg1, srcRowIndex, srcColumnIndex);
                result = Evaluate(d0, d1);
            }
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }

        private static double Evaluate(DateTime endDate, DateTime startDate)
        {
            return (endDate-startDate).Days;
        }

        private static DateTime getDate(ValueEval eval, int srcRowIndex, int srcColumnIndex)
        {
            ValueEval ve = OperandResolver.GetSingleValue(eval, srcRowIndex, srcColumnIndex);
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(ve, srcRowIndex, srcColumnIndex);
                return getDate(d0);
            }
            catch(Exception e)
            {
                String strText1 = OperandResolver.CoerceValueToString(ve);
                DateTime result = DateTime.MinValue;
                if(DateTime.TryParse(strText1,out result))
                    return result;

                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
        }

        private static DateTime getDate(double date)
        {
            DateTime d = DateUtil.GetJavaDate(date, false);
            return d.ToLocalTime();
        }
    }
}

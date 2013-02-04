using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;

namespace NPOI.SS.Formula.Functions
{
    public class Days360 : Var2or3ArgFunction
    {

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                result = Evaluate(d0, d1);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1, ValueEval arg2)
        {
            double result;
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                ValueEval ve = OperandResolver.GetSingleValue(arg2, srcRowIndex, srcColumnIndex);
                bool? method = OperandResolver.CoerceValueToBoolean(ve, false);
                result = Evaluate(d0, d1);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        private double Evaluate(double d0, double d1)
        {
            DateTime startingDate = GetStartingDate(d0);
            DateTime endingDate = GetEndingDateAccordingToStartingDate(d1, startingDate);
            long startingDay = startingDate.Month * 30 + startingDate.Day;
            long endingDay = (endingDate.Year - startingDate.Year) * 360
                    + endingDate.Month * 30 + endingDate.Day;
            return endingDay - startingDay;
        }
        private DateTime GetDate(double date)
        {
            return DateUtil.GetJavaDate(date);
        }
        private DateTime GetStartingDate(double date)
        {
            DateTime startingDate = GetDate(date);
            if (IsLastDayOfMonth(startingDate))
            {
                startingDate = new DateTime(startingDate.Year, startingDate.Month, 30, startingDate.Hour, startingDate.Minute, startingDate.Second);
            }
            return startingDate;
        }
        private DateTime GetEndingDateAccordingToStartingDate(double date, DateTime startingDate)
        {
            DateTime endingDate = DateUtil.GetJavaDate(date, false);
            if (IsLastDayOfMonth(endingDate))
            {
                if (startingDate.Day < 30)
                {
                    endingDate = GetFirstDayOfNextMonth(endingDate);
                }
            }
            return endingDate;
        }
        private bool IsLastDayOfMonth(DateTime date)
        {
            return date.AddDays(1).Month != date.Month;
        }
        private DateTime GetFirstDayOfNextMonth(DateTime date)
        {
            DateTime newDate;
            if (date.Month < 12)
            {
                newDate = new DateTime(date.Year, date.Month + 1, 1, date.Hour, date.Minute, date.Second);
            }
            else
            {
                newDate = new DateTime(date.Year + 1, 1, 1, date.Hour, date.Minute, date.Second);
            }
            return newDate;
        }
    }
}
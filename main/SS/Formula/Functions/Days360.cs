using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;

namespace NPOI.SS.Formula.Functions
{
    /**
     * <p>Calculates the number of days between two dates based on a 360-day year
     * (twelve 30-day months), which is used in some accounting calculations. Use
     * this function to help compute payments if your accounting system is based on
     * twelve 30-day months.</p>
     * 
     * {@code DAYS360(start_date,end_date,[method])}
     * 
     * <ul>
     * <li>Start_date, end_date (required):<br/>
     * The two dates between which you want to know the number of days.<br/>
     * If start_date occurs after end_date, the DAYS360 function returns a negative number.</li>
     * 
     * <li>Method (optional):<br/>
     * A logical value that specifies whether to use the U.S. or European method in the calculation</li>
     * 
     * <li>Method set to false or omitted:<br/>
     * the DAYS360 function uses the U.S. (NASD) method. If the starting date is the 31st of a month,
     * it becomes equal to the 30th of the same month. If the ending date is the 31st of a month and
     * the starting date is earlier than the 30th of a month, the ending date becomes equal to the
     * 1st of the next month, otherwise the ending date becomes equal to the 30th of the same month.
     * The month February and leap years are handled in the following way:<br/>
     * On a non-leap year the function {@code =DAYS360("2/28/93", "3/1/93", FALSE)} returns 1 day
     * because the DAYS360 function ignores the extra days added to February.<br/>
     * On a leap year the function {@code =DAYS360("2/29/96","3/1/96", FALSE)} returns 1 day for
     * the same reason.</li>
     * 
     * <li>Method Set to true:<br/>
     * When you set the method parameter to TRUE, the DAYS360 function uses the European method.
     * Starting dates or ending dates that occur on the 31st of a month become equal to the 30th of
     * the same month. The month February and leap years are handled in the following way:<br/>
     * On a non-leap year the function {@code =DAYS360("2/28/93", "3/1/93", TRUE)} returns
     * 3 days because the DAYS360 function is counting the extra days added to February to give
     * February 30 days.<br/>
     * On a leap year the function {@code =DAYS360("2/29/96", "3/1/96", TRUE)} returns
     * 2 days for the same reason.</li>
     * </ul>
     * 
     * @see <a href="https://support.microsoft.com/en-us/kb/235575">DAYS360 Function Produces Different Values Depending on the Version of Excel</a>
     */
    public class Days360 : Var2or3ArgFunction
    {

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                return new NumberEval(Evaluate(d0, d1, false));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1, ValueEval arg2)
        {
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                ValueEval ve = OperandResolver.GetSingleValue(arg2, srcRowIndex, srcColumnIndex);
                bool? method = OperandResolver.CoerceValueToBoolean(ve, false);
                return new NumberEval(Evaluate(d0, d1, method != null && (bool)method));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        private double Evaluate(double d0, double d1, bool method)
        {
            DateTime realStart = GetDate(d0);
            DateTime realEnd = GetDate(d1);
            int[] startingDate = GetStartingDate(realStart, method);
            int[] endingDate = GetEndingDate(realEnd, startingDate, method);

            return
                (endingDate[0] * 360 + endingDate[1] * 30 + endingDate[2]) -
                (startingDate[0] * 360 + startingDate[1] * 30 + startingDate[2]);
        }
        private DateTime GetDate(double date)
        {
            return DateUtil.GetJavaDate(date);
        }
        private int[] GetStartingDate(DateTime realStart, bool method)
        {
            int yyyy = realStart.Year;
            int mm = realStart.Month;
            int dd = Math.Min(30, realStart.Day);
            
            if (!method&& IsLastDayOfMonth(realStart)) dd = 30;
            return new int[] { yyyy, mm, dd };
        }

        private static int[] GetEndingDate(DateTime realEnd, DateTime realStart, bool method)
        {
            DateTime d = realEnd;
            int yyyy = d.Year;
            int mm = d.Month;
            int dd = Math.Min(30, d.Day);

            if (method == false && realEnd.Day == 31)
            {
                if (realStart.Day < 30)
                {
                    d = new DateTime(d.Year, d.Month, 1);
                    d = d.AddMonths(1);
                    yyyy = d.Year;
                    mm = d.Month;
                    dd = 1;
                }
                else
                {
                    dd = 30;
                }
            }

            return new int[] { yyyy, mm, dd };
        }
        private static int[] GetEndingDate(DateTime realEnd, int[] startingDate, bool method)
        {
            int yyyy = realEnd.Year;
            int mm = realEnd.Month;
            int dd = Math.Min(30, realEnd.Day);

            if (!method && realEnd.Day == 31)
            {
                if (startingDate[2] < 30)
                {
                    yyyy = realEnd.Year;
                    mm = realEnd.Month+1;
                    dd = 1;
                }
                else
                {
                    dd = 30;
                }
            }
            return new int[] { yyyy, mm, dd };
        }
        private DateTime GetEndingDateAccordingToStartingDate(double date, DateTime startingDate, bool method)
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
            //int dayOfMonth = date.Day;
            //int lastDayOfMonth = GetFirstDayOfNextMonth(date).AddDays(-1).Day;// getActualMaximum(Calendar.DAY_OF_MONTH);
            //return (dayOfMonth == lastDayOfMonth);
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
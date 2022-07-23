using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Atp
{
    public class WorkdayIntlFunction: FreeRefFunction
    {
        public static FreeRefFunction instance = new WorkdayIntlFunction(ArgumentsEvaluator.instance);

        private ArgumentsEvaluator evaluator;

        private WorkdayIntlFunction(ArgumentsEvaluator anEvaluator)
        {
            // enforces singleton
            this.evaluator = anEvaluator;
        }
        /**
         * Evaluate for WORKDAY. Given a date, a number of days and a optional date or interval of holidays, determines which date it is past
         * number of parametrized workdays.
         * 
         * @return {@link ValueEval} with date as its value.
         */
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 2 || args.Length > 4)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int srcCellRow = ec.RowIndex;
            int srcCellCol = ec.ColumnIndex;

            double start;
            int days;
            int weekendType = 1;
            double[] holidays;
            try
            {
                start = this.evaluator.EvaluateDateArg(args[0], srcCellRow, srcCellCol);
                days = (int)Math.Floor(this.evaluator.EvaluateNumberArg(args[1], srcCellRow, srcCellCol));
                if (args.Length >= 3)
                {
                    weekendType = (int)this.evaluator.EvaluateNumberArg(args[2], srcCellRow, srcCellCol);
                    if (!WorkdayCalculator.instance.GetValidWeekendTypes().Contains(weekendType))
                    {
                        return ErrorEval.NUM_ERROR;
                    }
                }
                ValueEval holidaysCell = args.Length>=4 ? args[3] : null;
                holidays = this.evaluator.EvaluateDatesArg(holidaysCell, srcCellRow, srcCellCol);
                return new NumberEval(DateUtil.GetExcelDate(
                    WorkdayCalculator.instance.CalculateWorkdays(start, days, weekendType, holidays)));
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }
        }
    }
}

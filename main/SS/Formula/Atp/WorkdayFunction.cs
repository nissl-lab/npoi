/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace NPOI.SS.Formula.Atp
{
    /**
 * Implementation of Excel 'Analysis ToolPak' function WORKDAY()<br/>
 * Returns the date past a number of workdays beginning at a start date, considering an interval of holidays. A workday is any non
 * saturday/sunday date.
 * <p/>
 * <b>Syntax</b><br/>
 * <b>WORKDAY</b>(<b>startDate</b>, <b>days</b>, holidays)
 * <p/>
 * 
 * @author jfaenomoto@gmail.com
 */
    class WorkdayFunction : FreeRefFunction
    {

        public static FreeRefFunction instance = new WorkdayFunction(ArgumentsEvaluator.instance);

        private ArgumentsEvaluator evaluator;

        private WorkdayFunction(ArgumentsEvaluator anEvaluator)
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
            if (args.Length < 2 || args.Length > 3)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int srcCellRow = ec.RowIndex;
            int srcCellCol = ec.ColumnIndex;

            double start;
            int days;
            double[] holidays;
            try
            {
                start = this.evaluator.EvaluateDateArg(args[0], srcCellRow, srcCellCol);
                days = (int)Math.Floor(this.evaluator.EvaluateNumberArg(args[1], srcCellRow, srcCellCol));
                ValueEval holidaysCell = args.Length == 3 ? args[2] : null;
                holidays = this.evaluator.EvaluateDatesArg(holidaysCell, srcCellRow, srcCellCol);
                return new NumberEval(DateUtil.GetExcelDate(WorkdayCalculator.instance.CalculateWorkdays(start, days, holidays)));
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }
        }

    }

}
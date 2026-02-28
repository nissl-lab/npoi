/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;

    /**
     * Implementation for the Excel EOMONTH() function.<p/>
     * <p/>
     * EOMONTH() returns the date of the last day of a month..<p/>
     * <p/>
     * <b>Syntax</b>:<br/>
     * <b>EOMONTH</b>(<b>start_date</b>,<b>months</b>)<p/>
     * <p/>
     * <b>start_date</b> is the starting date of the calculation
     * <b>months</b> is the number of months to be Added to <b>start_date</b>,
     * to give a new date. For this new date, <b>EOMONTH</b> returns the date of
     * the last day of the month. <b>months</b> may be positive (in the future),
     * zero or negative (in the past).
     */
    public class EOMonth : FreeRefFunction
    {

        public static FreeRefFunction instance = new EOMonth();


        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                double startDateAsNumber = NumericFunction.SingleOperandEvaluate(args[0], ec.RowIndex, ec.ColumnIndex);
                int months = (int)NumericFunction.SingleOperandEvaluate(args[1], ec.RowIndex, ec.ColumnIndex);

                // Excel treats date 0 as 1900-01-00; EOMONTH results in 1900-01-31
                if (startDateAsNumber >= 0.0 && startDateAsNumber < 1.0)
                {
                    startDateAsNumber = 1.0;
                }

                DateTime startDate = DateUtil.GetJavaDate(startDateAsNumber, false);
                //the month
                DateTime dtEnd = startDate.AddMonths(months);
                //next month
                dtEnd = dtEnd.AddMonths(1);
                //first day of next month
                dtEnd = new DateTime(dtEnd.Year, dtEnd.Month, 1);
                //last day of the month
                dtEnd = dtEnd.AddDays(-1);
                

                return new NumberEval(DateUtil.GetExcelDate(dtEnd));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

    }

}
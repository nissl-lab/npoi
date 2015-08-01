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

using NPOI.SS.Formula.Eval;
using System;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using System.Globalization;
namespace NPOI.SS.Formula.Functions
{



    /**
     * Implementation for Excel WeekNum() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>WeekNum  </b>(<b>Serial_num</b>,<b>Return_type</b>)<br/>
     * <p/>
     * Returns a number that indicates where the week falls numerically within a year.
     * <p/>
     * <p/>
     * Serial_num     is a date within the week. Dates should be entered by using the DATE function,
     * or as results of other formulas or functions. For example, use DATE(2008,5,23)
     * for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
     * Return_type     is a number that determines on which day the week begins. The default is 1.
     * 1	Week begins on Sunday. Weekdays are numbered 1 through 7.
     * 2	Week begins on Monday. Weekdays are numbered 1 through 7.
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class WeekNum : Fixed2ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new WeekNum();

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval serialNumVE, ValueEval returnTypeVE)
        {
            double serialNum = 0.0;
            try
            {
                serialNum = NumericFunction.SingleOperandEvaluate(serialNumVE, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }
            //Calendar serialNumCalendar = new GregorianCalendar();
            //serialNumCalendar.setTime(DateUtil.GetJavaDate(serialNum, false));
            DateTime serialNumCalendar = DateUtil.GetJavaDate(serialNum, false);

            int returnType = 0;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(returnTypeVE, srcRowIndex, srcColumnIndex);
                returnType = OperandResolver.CoerceValueToInt(ve);
            }
            catch (EvaluationException)
            {
                return ErrorEval.NUM_ERROR;
            }

            if (returnType != 1 && returnType != 2)
            {
                return ErrorEval.NUM_ERROR;
            }

            return new NumberEval(this.getWeekNo(serialNumCalendar, returnType));
        }

        public int getWeekNo(DateTime dt, int weekStartOn)
        {
            GregorianCalendar cal = new GregorianCalendar();
            int weekOfYear;
            if (weekStartOn == 1)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
            }
            else
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            }
            return weekOfYear;
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 2)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
            }
            return ErrorEval.VALUE_INVALID;
        }
    }
}
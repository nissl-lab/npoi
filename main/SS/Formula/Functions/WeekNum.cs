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
using System.Collections.Generic;

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
        private static NumberEval DEFAULT_RETURN_TYPE = new NumberEval(1);
        private static List<int> VALID_RETURN_TYPES = new List<int>() { 1, 2, 11, 12, 13, 14, 15, 16, 17, 21 };

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval serialNumVE, ValueEval returnTypeVE)
        {
            double serialNum;
            try
            {
                serialNum = NumericFunction.SingleOperandEvaluate(serialNumVE, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }
            DateTime serialNumCalendar;
            try
            {
                serialNumCalendar = DateUtil.GetJavaDate(serialNum, false);
            }
            catch (Exception )
            {
                return ErrorEval.NUM_ERROR;
            }
            int returnType;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(returnTypeVE, srcRowIndex, srcColumnIndex);
                returnType = OperandResolver.CoerceValueToInt(ve);
                if (ve is MissingArgEval)
                {
                    returnType = (int)DEFAULT_RETURN_TYPE.NumberValue;
                }
                else {
                    returnType = OperandResolver.CoerceValueToInt(ve);
                }
            }
            catch (EvaluationException)
            {
                return ErrorEval.NUM_ERROR;
            }

            if (!VALID_RETURN_TYPES.Contains(returnType))
            {
                return ErrorEval.NUM_ERROR;
            }

            return new NumberEval(this.getWeekNo(serialNumCalendar, returnType));
        }


        public int getWeekNo(DateTime dt, int weekStartOn)
        {
            GregorianCalendar cal = new GregorianCalendar();
            int weekOfYear;
            if (weekStartOn == 1 || weekStartOn == 17)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
            }
            else if (weekStartOn == 2 || weekStartOn == 11)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            }
            else if (weekStartOn == 12)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Tuesday);
            }
            else if (weekStartOn == 13)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Wednesday);
            }
            else if (weekStartOn == 14)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Thursday);
            }
            else if (weekStartOn == 15)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Friday);
            }
            else if (weekStartOn == 16)
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Saturday);
            }
            else
            {
                weekOfYear = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            }
            return weekOfYear;
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 1)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], DEFAULT_RETURN_TYPE);
            }
            else if (args.Length == 2)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
            }
            return ErrorEval.VALUE_INVALID;
        }
    }
}
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
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using System;

    /**
     * Implementation of Excel functions Date parsing functions:
     *  Date - DAY, MONTH and YEAR
     *  Time - HOUR, MINUTE and SECOND
     *  
     * @author Others (not mentioned in code)
     * @author Thies Wellpott
     */
    public class CalendarFieldFunction : Fixed1ArgFunction
    {
        public const int YEAR_ID = 0x01;
        public const int MONTH_ID = 0x02;
        public const int DAY_OF_MONTH_ID = 0x03;
        public const int HOUR_OF_DAY_ID = 0x04;
        public const int MINUTE_ID = 0x05;
        public const int SECOND_ID = 0x06;

        public static readonly Function YEAR = new CalendarFieldFunction(YEAR_ID);
        public static readonly Function MONTH = new CalendarFieldFunction(MONTH_ID);
        public static readonly Function DAY = new CalendarFieldFunction(DAY_OF_MONTH_ID);
        public static readonly Function HOUR = new CalendarFieldFunction(HOUR_OF_DAY_ID);
        public static readonly Function MINUTE = new CalendarFieldFunction(MINUTE_ID);
        public static readonly Function SECOND = new CalendarFieldFunction(SECOND_ID);

        private int _dateFieldId;

        private CalendarFieldFunction(int dateFieldId)
        {
            _dateFieldId = dateFieldId;
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            double val;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                val = OperandResolver.CoerceValueToDouble(ve);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            if (val < 0)
            {
                return ErrorEval.NUM_ERROR;
            }
            return new NumberEval(GetCalField(val));
        }

        private int GetCalField(double serialDate)
        {
            if ((int)serialDate == 0)
            {
                // Special weird case
                // day zero should be 31-Dec-1899,  but Excel seems to think it is 0-Jan-1900
                switch (_dateFieldId)
                {
                    case YEAR_ID: return 1900;
                    case MONTH_ID: return 1;
                    case DAY_OF_MONTH_ID: return 0;
                }
                //throw new InvalidOperationException("bad date field " + _dateFieldId);
            }
            // TODO Figure out if we're in 1900 or 1904
            // EXCEL functions round up nearly a half second (probably to prevent floating point
            // rounding issues); use UTC here to prevent daylight saving issues for HOUR
            //DateTime d = DateUtil.GetJavaDate(serialDate, false);
            DateTime d = DateUtil.GetJavaCalendar(serialDate + 0.4995 / DateUtil.SECONDS_PER_DAY, false);

            //Calendar c = new GregorianCalendar();
            //c.setTime(d);
            int result = 0;
            if (_dateFieldId == YEAR_ID)
            {
                result = d.Year;
            }
            else if (_dateFieldId == MONTH_ID)
            {
                result = d.Month;
            }
            else if (_dateFieldId == DAY_OF_MONTH_ID)
            {
                result = d.Day;
            }
            else if (_dateFieldId == HOUR_OF_DAY_ID)
            {
                result = d.Hour;
            }
            else if (_dateFieldId == MINUTE_ID)
            {
                result = d.Minute;
            }
            else if (_dateFieldId == SECOND_ID)
            {
                result = d.Second;
            }
            return result;
        }
    }
}


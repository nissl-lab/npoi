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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    

    //import java.util.Calendar;
    //import java.util.Date;
    //import java.util.GregorianCalendar;

    /**
     * Implementation of Excel functions DAY, MONTH and YEAR
     * 
     * 
     * @author Guenter Kickinger g.kickinger@gmx.net
     */
    public class CalendarFieldFunction : Function
    {
        public const int YEAR_ID=0x01;
        public const int MONTH_ID=0x02;
        public const int DAY_OF_MONTH_ID=0x03;

        public static Function YEAR = new CalendarFieldFunction(YEAR_ID, false);
        public static Function MONTH = new CalendarFieldFunction(MONTH_ID, false);
        public static Function DAY = new CalendarFieldFunction(DAY_OF_MONTH_ID, false);

        private int _dateFieldId;
        private bool _needsOneBaseAdjustment;

        private CalendarFieldFunction(int dateFieldId, bool needsOneBaseAdjustment)
        {
            _dateFieldId = dateFieldId;
            _needsOneBaseAdjustment = needsOneBaseAdjustment;
        }

        public ValueEval Evaluate(ValueEval[] operands, int srcCellRow, int srcCellCol)
        {
            if (operands.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int val;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(operands[0], srcCellRow, srcCellCol);
                val = OperandResolver.CoerceValueToInt(ve);
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

        private int GetCalField(int serialDay)
        {
            if (serialDay == 0)
            {
                // Special weird case
                // day zero should be 31-Dec-1899,  but Excel seems to think it is 0-Jan-1900
                switch (_dateFieldId)
                {
                    case YEAR_ID: return 1900;
                    case MONTH_ID: return 1;
                    case DAY_OF_MONTH_ID: return 0;
                }
                throw new InvalidOperationException("bad date field " + _dateFieldId);
            }
            DateTime d = NPOI.SS.UserModel.DateUtil.GetJavaDate(serialDay, false); // TODO fix 1900/1904 problem

            //Calendar c = new GregorianCalendar();
            //c.setTime(d);
            int result=0;
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
            if (_needsOneBaseAdjustment)
            {
                result++;
            }
            return result;
        }
    }
}
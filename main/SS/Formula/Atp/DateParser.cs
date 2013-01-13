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
using NPOI.Util;

namespace NPOI.SS.Formula.Atp
{
    /**
     * Parser for java dates.
     * 
     * @author jfaenomoto@gmail.com
     */
    public class DateParser
    {

        public DateParser instance = new DateParser();

        private DateParser()
        {
            // enforcing singleton
        }

        /**
         * Parses a date from a string.
         * 
         * @param strVal a string with a date pattern.
         * @return a date parsed from argument.
         * @throws EvaluationException exception upon parsing.
         */
        public static DateTime ParseDate(String strVal)
        {
            String[] parts = strVal.Split("-/".ToCharArray());// Pattern.compile("/").split(strVal);
            if (parts.Length != 3)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            String part2 = parts[2];
            int spacePos = part2.IndexOf(' ');
            if (spacePos > 0)
            {
                // drop time portion if present
                part2 = part2.Substring(0, spacePos);
            }
            int f0;
            int f1;
            int f2;
            try
            {
                f0 = int.Parse(parts[0]);
                f1 = int.Parse(parts[1]);
                f2 = int.Parse(part2);
            }
            catch (FormatException)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            if (f0 < 0 || f1 < 0 || f2 < 0 || (f0 > 12 && f1 > 12 && f2 > 12))
            {
                // easy to see this cannot be a valid date
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            if (f0 >= 1900 && f0 < 9999)
            {
                // when 4 digit value appears first, the format is YYYY/MM/DD, regardless of OS settings
                return MakeDate(f0, f1, f2);
            }
#if !HIDE_UNREACHABLE_CODE
            // otherwise the format seems to depend on OS settings (default date format)
            if (false)
            {
                // MM/DD/YYYY is probably a good guess, if the in the US
                return MakeDate(f2, f0, f1);
            }
#endif
            // TODO - find a way to choose the correct date format
            throw new RuntimeException("Unable to determine date format for text '" + strVal + "'");
        }

        /**
         * @param month 1-based
         */
        private static DateTime MakeDate(int year, int month, int day)
        {
            if (month < 1 || month > 12)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            DateTime cal = new DateTime(year, month, 1, 0, 0, 0);
            //cal.set(Calendar.MILLISECOND, 0);
            int maxDay = cal.AddMonths(1).AddDays(-1).Day;
            if (day < 1 || day > maxDay /*cal.GetActualMaximum(Calendar.DAY_OF_MONTH)*/)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            //cal.set(Calendar.DAY_OF_MONTH, day);
            return new DateTime(year, month, day);
        }

    }

}
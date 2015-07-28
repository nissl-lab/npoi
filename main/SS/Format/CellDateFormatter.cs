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

using System;
using System.Text.RegularExpressions;
using System.Text;
using NPOI.SS.Util;
using System.Globalization;

namespace NPOI.SS.Format
{
    /**
     * Formats a date value.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public class CellDateFormatter : CellFormatter
    {
        private bool amPmUpper;
        private bool ShowM;
        private bool ShowAmPm;
        private FormatBase dateFmt;
        private String sFmt;
        private int millisecondPartLength = 0;

        private static readonly TimeSpan EXCEL_EPOCH_TIME;
        private static readonly DateTime EXCEL_EPOCH_DATE;

        private static readonly CellFormatter SIMPLE_DATE = new CellDateFormatter(
                "mm/d/y");

        static CellDateFormatter()
        {
            DateTime c = new DateTime(1904, 1, 1);
            EXCEL_EPOCH_DATE = c;
            EXCEL_EPOCH_TIME = c.TimeOfDay;
        }

        private class DatePartHandler : CellFormatPart.IPartHandler
        {
            private CellDateFormatter _formatter;
            private int mStart = -1;
            private int mLen;
            private int hStart = -1;
            private int hLen;
            public DatePartHandler(CellDateFormatter formatter)
            {
                this._formatter = formatter;
            }
            public String HandlePart(Match m, String part, CellFormatType type, StringBuilder desc)
            {

                int pos = desc.Length;
                char firstCh = part[0];
                switch (firstCh)
                {
                    case 's':
                    case 'S':
                        if (mStart >= 0)
                        {
                            for (int i = 0; i < mLen; i++)
                                desc[mStart + i] = 'm';
                            mStart = -1;
                        }
                        return part.ToLower();

                    case 'h':
                    case 'H':
                        mStart = -1;
                        hStart = pos;
                        hLen = part.Length;
                        return part.ToLower();

                    case 'd':
                    case 'D':
                        mStart = -1;
                        //if (part.Length <= 2)
                            return part.ToLower();
                        //else
                        //    return part.ToLower().Replace('d', 'E');

                    case 'm':
                    case 'M':
                        mStart = pos;
                        mLen = part.Length;
                        // For 'm' after 'h', output minutes ('m') not month ('M')
                        if (hStart >= 0)
                            return part.ToLower();
                        else
                            return part.ToUpper();

                    case 'y':
                    case 'Y':
                        mStart = -1;
                        if (part.Length == 3)
                            part = "yyyy";
                        return part.ToLower();

                    case '0':
                        mStart = -1;
                        int sLen = part.Length;
                        _formatter.sFmt = "%0" + (sLen + 2) + "." + sLen + "f";
                        _formatter.millisecondPartLength = sLen;
                        return part.Replace('0', 'f');

                    case 'a':
                    case 'A':
                    case 'p':
                    case 'P':
                        if (part.Length > 1)
                        {
                            // am/pm marker
                            mStart = -1;
                            _formatter.ShowAmPm = true;
                            _formatter.ShowM = char.ToLower(part[1]) == 'm';
                            // For some reason "am/pm" becomes AM or PM, but "a/p" becomes a or p
                            _formatter.amPmUpper = _formatter.ShowM || char.IsUpper(part[0]);
                            if (_formatter.ShowM)
                                return "tt";
                            else
                                return "t";
                            //return "a";
                        }
                    //noinspection fallthrough
                        return null;
                    default:
                        return null;
                }
            }

            public void Finish(StringBuilder toAppendTo)
            {
                if (hStart >= 0 && !_formatter.ShowAmPm)
                {
                    for (int i = 0; i < hLen; i++)
                    {
                        toAppendTo[hStart + i] = 'H';
                    }
                }
            }
        }

        /**
         * Creates a new date formatter with the given specification.
         *
         * @param format The format.
         */
        public CellDateFormatter(String format)
            : base(format)
        {
            DatePartHandler partHandler = new DatePartHandler(this);
            StringBuilder descBuf = CellFormatPart.ParseFormat(format,
                    CellFormatType.DATE, partHandler);
            partHandler.Finish(descBuf);
            dateFmt = new SimpleDateFormat(descBuf.ToString());
            
            // tweak the format pattern to pass tests on JDK 1.7,
            // See https://issues.apache.org/bugzilla/show_bug.cgi?id=53369

            String ptrn = Regex.Replace(descBuf.ToString(), "((y)(?!y))(?<!yy)", "yy");
            dateFmt = new SimpleDateFormat(ptrn);
        }

        /** {@inheritDoc} */
        public override void FormatValue(StringBuilder toAppendTo, Object value)
        {
            if (value == null)
                value = 0.0;
            //if (value is Number) {
            if (NPOI.Util.Number.IsNumber(value))
            {
                double v;
                double.TryParse(value.ToString(), out v);
                if (v == 0.0)
                    value = EXCEL_EPOCH_DATE;
                else
                    value = new DateTime((long)(EXCEL_EPOCH_TIME.Ticks + v));
            }
            DateTime newValue = (DateTime)value;
            //in excel, if  millisecond part value like .009, (that is above 5 millisecond)
            //and millisecond part pattern is .00 (changed to .ff), the result need round up,
            //millisecond part is 01 not 00. so try to adjuse the value.
            if (millisecondPartLength > 0)
            {
                double second = (newValue.Millisecond / 1000.0 * Math.Pow(10, millisecondPartLength));
                second = second - Math.Truncate(second);
                if (second >= 0.5)
                {
                    newValue = newValue.AddMilliseconds((1 - second) / Math.Pow(10, millisecondPartLength) * 1000);
                }
            }

            dateFmt.Format(newValue, toAppendTo, CultureInfo.CurrentCulture);
            
            //throw new NotImplementedException();
            //AttributedCharacterIterator it = dateFmt.FormatToCharacterIterator(
            //        value);
            //bool doneAm = false;
            //bool doneMillis = false;

            //it.First();
            //for (char ch = it.First();
            //     ch != CharacterIterator.DONE;
            //     ch = it.Next())
            //{
            //    if (it.GetAttribute(DateFormat.Field.MILLISECOND) != null)
            //    {
            //        if (!doneMillis)
            //        {
            //            Date dateObj = (Date)value;
            //            int pos = toAppendTo.Length();
            //            Formatter formatter = new Formatter(toAppendTo);
            //            long msecs = dateObj.Time % 1000;
            //            formatter.Format(LOCALE, sFmt, msecs / 1000.0);
            //            toAppendTo.Remove(pos, 2);
            //            doneMillis = true;
            //        }
            //    }
            //    else if (it.GetAttribute(DateFormat.Field.AM_PM) != null)
            //    {
            //        if (!doneAm)
            //        {
            //            if (ShowAmPm)
            //            {
            //                if (amPmUpper)
            //                {
            //                    toAppendTo.Append(char.ToUpper(ch));
            //                    if (ShowM)
            //                        toAppendTo.Append('M');
            //                }
            //                else
            //                {
            //                    toAppendTo.Append(char.ToLower(ch));
            //                    if (ShowM)
            //                        toAppendTo.Append('m');
            //                }
            //            }
            //            doneAm = true;
            //        }
            //    }
            //    else
            //    {
            //        toAppendTo.Append(ch);
            //    }
            //}
        }

        /**
         * {@inheritDoc}
         * <p/>
         * For a date, this is <tt>"mm/d/y"</tt>.
         */
        public override void SimpleValue(StringBuilder toAppendTo, Object value)
        {
            SIMPLE_DATE.FormatValue(toAppendTo, value);
        }
    }
}
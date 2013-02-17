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
namespace NPOI.SS.Format
{
    using System;
    using System.Text;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    /**
     * This class : printing out an elapsed time format.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public class CellElapsedFormatter : CellFormatter
    {
        private List<TimeSpec> specs;
        private TimeSpec topmost;
        private String printfFmt;

        private static readonly Regex PERCENTS = new Regex("%");

        private const double HOUR__FACTOR = 1.0 / 24.0;
        private const double MIN__FACTOR = HOUR__FACTOR / 60.0;
        private const double SEC__FACTOR = MIN__FACTOR / 60.0;

        private class TimeSpec
        {
            internal char type;
            internal int pos;
            internal int len;
            internal double factor;
            internal double modBy;

            public TimeSpec(char type, int pos, int len, double factor)
            {
                this.type = type;
                this.pos = pos;
                this.len = len;
                this.factor = factor;
                modBy = 0;
            }

            public long ValueFor(double elapsed)
            {
                double val;
                if (modBy == 0)
                    val = elapsed / factor;
                else
                    val = elapsed / factor % modBy;
                if (type == '0')
                    return (long)Math.Round(val);
                else
                    return (long)val;
            }
        }

        private class ElapsedPartHandler : CellFormatPart.IPartHandler
        {
            public ElapsedPartHandler(CellElapsedFormatter formatter)
            {
                this._formatter = formatter;
            }
            // This is the one class that's directly using printf, so it can't use
            // the default handling for quoted strings and special characters.  The
            // only special character for this is '%', so we have to handle all the
            // quoting in this method ourselves.
            private CellElapsedFormatter _formatter;
            public String HandlePart(Match m, String part, CellFormatType type,
                    StringBuilder desc)
            {

                int pos = desc.Length;
                char firstCh = part[0];
                switch (firstCh)
                {
                    case '[':
                        if (part.Length < 3)
                            break;
                        if (_formatter.topmost != null)
                            throw new ArgumentException(
                                    "Duplicate '[' times in format");
                        part = part.ToLower();
                        int specLen = part.Length - 2;
                        _formatter.topmost = _formatter.AssignSpec(part[1], pos, specLen);
                        return part.Substring(1, specLen);

                    case 'h':
                    case 'm':
                    case 's':
                    case '0':
                        part = part.ToLower();
                        _formatter.AssignSpec(part[0], pos, part.Length);
                        return part;

                    case '\n':
                        return "%n";

                    case '\"':
                        part = part.Substring(1, part.Length - 2);
                        break;

                    case '\\':
                        part = part.Substring(1);
                        break;

                    case '*':
                        if (part.Length > 1)
                            part = CellFormatPart.ExpandChar(part);
                        break;

                    // An escape we can let it handle because it can't have a '%'
                    case '_':
                        return null;
                }
                // Replace ever "%" with a "%%" so we can use printf
                //return PERCENTS.Matcher(part).ReplaceAll("%%");
                //return PERCENTS.Replace(part, "%%");
                return part;
            }

        }

        /**
         * Creates a elapsed time formatter.
         *
         * @param pattern The pattern to Parse.
         */
        public CellElapsedFormatter(String pattern)
            : base(pattern)
        {
            specs = new List<TimeSpec>();

            StringBuilder desc = CellFormatPart.ParseFormat(pattern,
                    CellFormatType.ELAPSED, new ElapsedPartHandler(this));

            //ListIterator<TimeSpec> it = specs.ListIterator(specs.Count);
            //while (it.HasPrevious())
            for(int i=specs.Count-1;i>=0;i--)
            {
                //TimeSpec spec = it.Previous();
                TimeSpec spec = specs[i];
                //desc.Replace(spec.pos, spec.pos + spec.len, "%0" + spec.len + "d");
                desc.Remove(spec.pos, spec.len);
                desc.Insert(spec.pos, "D" + spec.len);
                if (spec.type != topmost.type)
                {
                    spec.modBy = modFor(spec.type, spec.len);
                }
            }

            printfFmt = desc.ToString();
        }

        private TimeSpec AssignSpec(char type, int pos, int len)
        {
            TimeSpec spec = new TimeSpec(type, pos, len, factorFor(type, len));
            specs.Add(spec);
            return spec;
        }

        private static double factorFor(char type, int len)
        {
            switch (type)
            {
                case 'h':
                    return HOUR__FACTOR;
                case 'm':
                    return MIN__FACTOR;
                case 's':
                    return SEC__FACTOR;
                case '0':
                    return SEC__FACTOR / Math.Pow(10, len);
                default:
                    throw new ArgumentException(
                            "Uknown elapsed time spec: " + type);
            }
        }

        private static double modFor(char type, int len)
        {
            switch (type)
            {
                case 'h':
                    return 24;
                case 'm':
                    return 60;
                case 's':
                    return 60;
                case '0':
                    return Math.Pow(10, len);
                default:
                    throw new ArgumentException(
                            "Uknown elapsed time spec: " + type);
            }
        }

        /** {@inheritDoc} */
        public override void FormatValue(StringBuilder toAppendTo, Object value)
        {
            double elapsed = ((double)value);

            if (elapsed < 0)
            {
                toAppendTo.Append('-');
                elapsed = -elapsed;
            }

            long[] parts = new long[specs.Count];
            
            for (int i = 0; i < specs.Count; i++)
            {
                parts[i] = specs[(i)].ValueFor(elapsed);
            }

            //Formatter formatter = new Formatter(toAppendTo);
            //formatter.Format(printfFmt, parts);
            string[] fmtPart = printfFmt.Split(":. []".ToCharArray());
            string split = string.Empty;
            int pos = 0;
            int index = 0;
            Regex regFmt = new Regex("D\\d+");
            foreach (string fmt in fmtPart)
            {
                pos += fmt.Length;
                if (pos < printfFmt.Length)
                {
                    split = printfFmt[pos].ToString();
                    pos++;
                }
                else
                    split = string.Empty;
                if (regFmt.IsMatch(fmt))
                {
                    toAppendTo.Append(parts[index].ToString(fmt)).Append(split);
                    index++;
                }
                else
                {
                    toAppendTo.Append(fmt).Append(split);
                }
            }

            //throw new NotImplementedException();
        }

        /**
         * {@inheritDoc}
         * <p/>
         * For a date, this is <tt>"mm/d/y"</tt>.
         */
        public override void SimpleValue(StringBuilder toAppendTo, Object value)
        {
            FormatValue(toAppendTo, value);
        }
    }
}
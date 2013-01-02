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



    /**
     * A formatter for the default "General" cell format.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public class CellGeneralFormatter : CellFormatter
    {
        /** Creates a new general formatter. */
        public CellGeneralFormatter()
            : base("General")
        {
            ;
        }

        /**
         * The general style is not quite the same as any other, or any combination
         * of others.
         *
         * @param toAppendTo The buffer to append to.
         * @param value      The value to format.
         */
        public override void FormatValue(StringBuilder toAppendTo, Object value)
        {
            //if (value is Number) {
            if (NPOI.Util.Number.IsNumber(value))
            {
                double val ;
                double.TryParse(value.ToString(), out val);
                if (val == 0)
                {
                    toAppendTo.Append('0');
                    return;
                }

                String fmt;
                double exp = Math.Log10(Math.Abs(val));
                bool stripZeros = true;
                if (exp > 10 || exp < -9)
                    fmt = "E5";
                else if ((long)val != val)
                    fmt = "F9";
                else
                {
                    fmt = "F0";
                    stripZeros = false;
                }
                toAppendTo.Append(val.ToString(fmt));

                //Formatter formatter = new Formatter(toAppendTo);
                //formatter.Format(LOCALE, fmt, value);
                if (stripZeros)
                {
                    // strip off trailing zeros
                    int RemoveFrom;
                    if (fmt.StartsWith("E"))
                        RemoveFrom = toAppendTo.ToString().LastIndexOf("E") - 1;
                    else
                        RemoveFrom = toAppendTo.Length - 1;
                    while (toAppendTo[RemoveFrom] == '0')
                    {
                        toAppendTo.Remove(RemoveFrom--, 1);
                    }
                    if (toAppendTo[RemoveFrom] == '.')
                    {
                        toAppendTo.Remove(RemoveFrom--, 1);
                    }
                    // remove zeros after E   by antony.liu
                    string text = toAppendTo.ToString();
                    RemoveFrom = toAppendTo.ToString().LastIndexOf("E");
                    if (RemoveFrom > 0)
                    {
                        RemoveFrom++;
                        if (text[RemoveFrom] == '+' || text[RemoveFrom] == '-')
                            RemoveFrom++;
                        int count = 0;
                        while (RemoveFrom + count< text.Length)
                        {
                            if (text[RemoveFrom + count] == '0')
                                count++;
                            else
                                break;
                        }
                        toAppendTo.Remove(RemoveFrom, count);
                    }
                }
            } 
            else if (value is Boolean) {
                toAppendTo.Append(value.ToString().ToUpper());
            }
            else
            {
                toAppendTo.Append(value.ToString());
            }
        }

        /** Equivalent to {@link #formatValue(StringBuilder,Object)}. {@inheritDoc}. */
        public override void SimpleValue(StringBuilder toAppendTo, Object value)
        {
            FormatValue(toAppendTo, value);
        }
    }
}
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
            if (value.GetType().IsPrimitive)
            {
                double val = ((double)value);
                if (val == 0)
                {
                    toAppendTo.Append('0');
                    return;
                }

                String fmt;
                double exp = Math.Log10(Math.Abs(val));
                bool stripZeros = true;
                if (exp > 10 || exp < -9)
                    fmt = "%1.5E";
                else if ((long)val != val)
                    fmt = "%1.9f";
                else
                {
                    fmt = "%1.0f";
                    stripZeros = false;
                }
                toAppendTo.Append(val.ToString(fmt));
                //Todo: Remove trailing zeros         commented by Antony.liu


                //Formatter formatter = new Formatter(toAppendTo);
                //formatter.Format(LOCALE, fmt, value);
                if (stripZeros)
                {
                //    // strip off trailing zeros
                //    int RemoveFrom;
                //    if (fmt.EndsWith("E"))
                //        RemoveFrom = toAppendTo.LastIndexOf("E") - 1;
                //    else
                //        RemoveFrom = toAppendTo.Length - 1;
                //    while (toAppendTo[RemoveFrom] == '0')
                //    {
                //        toAppendTo.DeleteCharAt(RemoveFrom--);
                //    }
                //    if (toAppendTo[RemoveFrom] == '.')
                //    {
                //        toAppendTo.DeleteCharAt(RemoveFrom--);
                //    }
                }
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
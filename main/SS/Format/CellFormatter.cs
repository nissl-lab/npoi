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
    using System.Globalization;




    /**
     * This is the abstract supertype for the various cell formatters.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public abstract class CellFormatter
    {
        /** The original specified format. */
        protected String format;

        /**
         * This is the locale used to Get a consistent format result from which to
         * work.
         */
        public static readonly CultureInfo LOCALE = CultureInfo.GetCultureInfo("en-US");

        /**
         * Creates a new formatter object, storing the format in {@link #format}.
         *
         * @param format The format.
         */
        public CellFormatter(String format)
        {
            this.format = format;
        }

        // /** The logger to use in the formatting code. */
        //static Logger logger = Logger.GetLogger(typeof(CellFormatter).Name);

        /**
         * Format a value according the format string.
         *
         * @param toAppendTo The buffer to append to.
         * @param value      The value to format.
         */
        public abstract void FormatValue(StringBuilder toAppendTo, Object value);

        /**
         * Format a value according to the type, in the most basic way.
         *
         * @param toAppendTo The buffer to append to.
         * @param value      The value to format.
         */
        public abstract void SimpleValue(StringBuilder toAppendTo, Object value);

        /**
         * Formats the value, returning the resulting string.
         *
         * @param value The value to format.
         *
         * @return The value, formatted.
         */
        public String Format(Object value)
        {
            StringBuilder sb = new StringBuilder();
            FormatValue(sb, value);
            return sb.ToString();
        }

        /**
         * Formats the value in the most basic way, returning the resulting string.
         *
         * @param value The value to format.
         *
         * @return The value, formatted.
         */
        public String SimpleFormat(Object value)
        {
            StringBuilder sb = new StringBuilder();
            SimpleValue(sb, value);
            return sb.ToString();
        }

        /**
         * Returns the input string, surrounded by quotes.
         *
         * @param str The string to quote.
         *
         * @return The input string, surrounded by quotes.
         */
        static String Quote(String str)
        {
            return '"' + str + '"';
        }

        public override string ToString()
        {
            return format;
        }
    }
}
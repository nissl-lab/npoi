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
    internal class GeneralCellFormatType : CellFormatType
    {
        public override CellFormatter Formatter(String pattern)
        {
            return new CellGeneralFormatter();
        }
        public override bool IsSpecial(char ch)
        {
            return false;
        }
    }
    internal class NumberCellFormatType : CellFormatType
    {
        public override CellFormatter Formatter(String pattern)
        {
            return new CellNumberFormatter(pattern);
        }
        public override bool IsSpecial(char ch)
        {
            return false;
        }
    }
    internal class DateCellFormatType : CellFormatType
    {
        public override bool IsSpecial(char ch)
        {
            return ch == '\'' || (ch <= '\u007f' && char.IsLetter(ch));
        }
        public override CellFormatter Formatter(String pattern)
        {
            return new CellDateFormatter(pattern);
        }
    }
    internal class ElapsedCellFormatType : CellFormatType
    {
        public override bool IsSpecial(char ch)
        {
            return false;
        }
        public override CellFormatter Formatter(String pattern)
        {
            return new CellElapsedFormatter(pattern);
        }
    }
    internal class TextCellFormatType : CellFormatType
    {
        public override bool IsSpecial(char ch)
        {
            return false;
        }
        public override CellFormatter Formatter(String pattern)
        {
            return new CellTextFormatter(pattern);
        }
    }
    /**
     * The different kinds of formats that the formatter understands.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public abstract class CellFormatType
    {
        /** The general (default) format; also used for <tt>"General"</tt>. */
        public static readonly CellFormatType GENERAL = new GeneralCellFormatType();
        /** A numeric format. */
        public static readonly CellFormatType NUMBER = new NumberCellFormatType();
        /** A date format. */
        public static readonly CellFormatType DATE = new DateCellFormatType();
        /** An elapsed time format. */
        public static readonly CellFormatType ELAPSED = new ElapsedCellFormatType();
        /** A text format. */
        public static readonly CellFormatType TEXT = new TextCellFormatType();

        /**
         * Returns <tt>true</tt> if the format is special and needs to be quoted.
         *
         * @param ch The character to test.
         *
         * @return <tt>true</tt> if the format is special and needs to be quoted.
         */
        public abstract bool IsSpecial(char ch);

        /**
         * Returns a new formatter of the appropriate type, for the given pattern.
         * The pattern must be appropriate for the type.
         *
         * @param pattern The pattern to use.
         *
         * @return A new formatter of the appropriate type, for the given pattern.
         */
        public abstract CellFormatter Formatter(String pattern);

    }

}
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

namespace NPOI.SS.UserModel
{
    /**
     * the different types of possible underline formatting
     *
     * @author Gisella Bronzetti
     */
    public class FontUnderline
    {

        /**
         * Single-line underlining under each character in the cell.
         * The underline is drawn through the descenders of
         * characters such as g and p..
         */
        public static readonly FontUnderline SINGLE = new FontUnderline(1);

        /**
         * Double-line underlining under each character in the
         * cell. underlines are drawn through the descenders of
         * characters such as g and p.
         */
        public static readonly FontUnderline DOUBLE = new FontUnderline(2);

        /**
         * Single-line accounting underlining under each
         * character in the cell. The underline is drawn under the
         * descenders of characters such as g and p.
         */
        public static readonly FontUnderline SINGLE_ACCOUNTING = new FontUnderline(3);

        /**
         * Double-line accounting underlining under each
         * character in the cell. The underlines are drawn under
         * the descenders of characters such as g and p.
         */
        public static readonly FontUnderline DOUBLE_ACCOUNTING = new FontUnderline(4);

        /**
         * No underline.
         */
        public static readonly FontUnderline NONE = new FontUnderline(0);


        private int value;


        private FontUnderline(int val)
        {
            value = val;
        }

        public int Value
        {
            get
            {
                return value;
            }
        }

        public byte ByteValue
        {
            get
            {
                if (this == DOUBLE)
                {
                    return (byte)FontUnderlineType.Double;
                }
                else if (this == DOUBLE_ACCOUNTING)
                {
                    return (byte)FontUnderlineType.DoubleAccounting;
                }
                else if (this == SINGLE_ACCOUNTING)
                {
                    return (byte)FontUnderlineType.SingleAccounting;
                }
                else if (this == NONE)
                {
                    return (byte)FontUnderlineType.None;
                }
                else if (this == SINGLE)
                {
                    return (byte)FontUnderlineType.Single;
                }
                else
                {
                    return (byte)FontUnderlineType.Single;
                }
            }
        }

        private static FontUnderline[] _table = null;

        static FontUnderline()
        {
            if (_table == null)
            {
                _table = new FontUnderline[5];
                _table[0] = FontUnderline.NONE;
                _table[1] = FontUnderline.SINGLE;
                _table[2] = FontUnderline.DOUBLE;
                _table[3] = FontUnderline.SINGLE_ACCOUNTING;
                _table[4] = FontUnderline.DOUBLE_ACCOUNTING;
            }
        }
        public static FontUnderline ValueOf(int value)
        {
            return _table[value];
        }

        public static FontUnderline ValueOf(FontUnderlineType value)
        {
            FontUnderline val;
            switch (value)
            {
                case FontUnderlineType.Double:
                    val = FontUnderline.DOUBLE;
                    break;
                case FontUnderlineType.DoubleAccounting:
                    val = FontUnderline.DOUBLE_ACCOUNTING;
                    break;
                case FontUnderlineType.SingleAccounting:
                    val = FontUnderline.SINGLE_ACCOUNTING;
                    break;
                case FontUnderlineType.Single:
                    val = FontUnderline.SINGLE;
                    break;
                default:
                    val = FontUnderline.NONE;
                    break;
            }
            return val;
        }

    }
}

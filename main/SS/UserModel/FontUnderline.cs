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
        public static readonly FontUnderline Single = new FontUnderline(1);

        /**
         * Double-line underlining under each character in the
         * cell. underlines are drawn through the descenders of
         * characters such as g and p.
         */
        public static readonly FontUnderline Double = new FontUnderline(2);

        /**
         * Single-line accounting underlining under each
         * character in the cell. The underline is drawn under the
         * descenders of characters such as g and p.
         */
        public static readonly FontUnderline SingleAccounting = new FontUnderline(3);

        /**
         * Double-line accounting underlining under each
         * character in the cell. The underlines are drawn under
         * the descenders of characters such as g and p.
         */
        public static readonly FontUnderline DoubleAccounting = new FontUnderline(4);

        /**
         * No underline.
         */
        public static readonly FontUnderline None = new FontUnderline(5);


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
                if (this == Double)
                {
                    return (byte)FontUnderlineType.Double;
                }
                else if (this == DoubleAccounting)
                {
                    return (byte)FontUnderlineType.DoubleAccounting;
                }
                else if (this == SingleAccounting)
                {
                    return (byte)FontUnderlineType.SingleAccounting;
                }
                else if (this == None)
                {
                    return (byte)FontUnderlineType.None;
                }
                else if (this == Single)
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
                _table = new FontUnderline[6];
                _table[1] = FontUnderline.Single;
                _table[2] = FontUnderline.Double;
                _table[3] = FontUnderline.SingleAccounting;
                _table[4] = FontUnderline.DoubleAccounting;
                _table[5] = FontUnderline.None;
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
                    val = FontUnderline.Double;
                    break;
                case FontUnderlineType.DoubleAccounting:
                    val = FontUnderline.DoubleAccounting;
                    break;
                case FontUnderlineType.SingleAccounting:
                    val = FontUnderline.SingleAccounting;
                    break;
                case FontUnderlineType.Single:
                    val = FontUnderline.Single;
                    break;
                default:
                    val = FontUnderline.None;
                    break;
            }
            return val;
        }

    }
}

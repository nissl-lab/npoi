/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.PTG
{
    using System;
    using System.Text;
    
    using NPOI.Util;


    /**
     * String Stores a String value in a formula value stored in the format
     * &lt;Length 2 bytes&gt;char[]
     * 
     * @author Werner Froidevaux
     * @author Jason Height (jheight at chariot dot net dot au)
     * @author Bernard Chesnoy
     */
    public class StringPtg : ScalarConstantPtg
    {
        public const byte sid = 0x17;
        private static BitField fHighByte = BitFieldFactory.GetInstance(0x01);
        /** the Char (")used in formulas to delimit string literals */
        private const char FORMULA_DELIMITER = '"';

        /**
         * NOTE: OO doc says 16bit Length, but BiffViewer says 8 Book says something
         * totally different, so don't look there!
         */
        private int field_1_Length;
        private byte field_2_options;      

        private bool _is16bitUnicode;
        private String field_3_string;

        /** Create a StringPtg from a stream */
        public StringPtg(ILittleEndianInput in1)
        {
            int field_1_length = in1.ReadUByte();
			field_2_options = (byte)in1.ReadByte();
            _is16bitUnicode = (field_2_options & 0x01) != 0;
            if (_is16bitUnicode)
            {
                field_3_string = StringUtil.ReadUnicodeLE(in1, field_1_length);
            }
            else
            {
                field_3_string = StringUtil.ReadCompressedUnicode(in1, field_1_length);
            }
        }

        /**
         * Create a StringPtg from a string representation of the number Number
         * format Is not Checked, it Is expected to be Validated in the Parser that
         * calls this method.
         * 
         * @param value :
         *            String representation of a floating point number
         */
        public StringPtg(String value)
        {
            if (value.Length > 255)
            {
                throw new ArgumentException(
                        "String literals in formulas can't be bigger than 255 Chars ASCII");
            }
            _is16bitUnicode = StringUtil.HasMultibyte(value);
            field_3_string = value;
            field_1_Length = value.Length; // for the moment, we support only ASCII strings in formulas we Create
        }

        public String Value
        {
            get { return field_3_string; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteByte(field_3_string.Length); // Note - nChars is 8-bit
            out1.WriteByte(_is16bitUnicode ? 0x01 : 0x00);
            if (_is16bitUnicode)
            {
                StringUtil.PutUnicodeLE(field_3_string, out1);
            }
            else
            {
                StringUtil.PutCompressedUnicode(field_3_string, out1);
            }
        }

        public override int Size
        {
            get
            {
                return field_3_string.Length * (_is16bitUnicode ? 2 : 1) + 3;
            }
        }

        public override String ToFormulaString()
        {
            String value = field_3_string;
            int len = value.Length;
            StringBuilder sb = new StringBuilder(len + 4);
            sb.Append(FORMULA_DELIMITER);

            for (int i = 0; i < len; i++)
            {
                char c = value[i];
                if (c == FORMULA_DELIMITER)
                {
                    sb.Append(FORMULA_DELIMITER);
                }
                sb.Append(c);
            }

            sb.Append(FORMULA_DELIMITER);
            return sb.ToString();
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(field_3_string);
            sb.Append("]");
            return sb.ToString();
        }
    }
}
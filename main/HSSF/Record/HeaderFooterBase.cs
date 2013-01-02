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

namespace NPOI.HSSF.Record
{
    using System;
using NPOI.Util;

/**
 * Common header/footer base class
 *
 * @author Josh Micich
 */
    public abstract class HeaderFooterBase : StandardRecord
    {
        private bool field_2_hasMultibyte;
        private String field_3_text;

        protected HeaderFooterBase(String text)
        {
            Text=(text);
        }

        protected HeaderFooterBase(RecordInputStream in1)
        {
            if (in1.Remaining > 0)
            {
                int field_1_footer_len = in1.ReadShort();
                field_2_hasMultibyte = in1.ReadByte() != 0x00;

                if (field_2_hasMultibyte)
                {
                    field_3_text = in1.ReadUnicodeLEString(field_1_footer_len);
                }
                else
                {
                    field_3_text = in1.ReadCompressedUnicode(field_1_footer_len);
                }
            }
            else
            {
                // Note - this is unusual for BIFF records in general, but normal for header / footer records:
                // when the text is empty string, the whole record is empty (just the 4 byte BIFF header)
                field_3_text = "";
            }
        }


        /**
         * get the length of the footer string
         *
         * @return length of the footer string
         */
        private int TextLength
        {
            get
            {
                return field_3_text.Length;
            }
        }

        public String Text
        {
            get
            {
                return field_3_text;
            }
            set 
            {
                if (value == null)
                {
                    throw new ArgumentException("text must not be null");
                }
                field_2_hasMultibyte = StringUtil.HasMultibyte(value);
                field_3_text = value;

                // Check it'll fit into the space in the record
                if (this.DataSize > RecordInputStream.MAX_RECORD_DATA_SIZE)
                {
                    throw new ArgumentException("Header/Footer string too long (limit is "
                            + RecordInputStream.MAX_RECORD_DATA_SIZE + " bytes)");
                }               
            }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            if (TextLength > 0)
            {
                out1.WriteShort(TextLength);
                out1.WriteByte(field_2_hasMultibyte ? 0x01 : 0x00);
                if (field_2_hasMultibyte)
                {
                    StringUtil.PutUnicodeLE(field_3_text, out1);
                }
                else
                {
                    StringUtil.PutCompressedUnicode(field_3_text, out1);
                }
            }
        }

        protected override int DataSize
        {
            get
            {
                if (TextLength< 1)
                {
                    return 0;
                }
                return 3 + TextLength * (field_2_hasMultibyte ? 2 : 1);
            }
        }
    }
}




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

    using NPOI.Util;

    using System.Text;
    using System;

    /**
     * Title: NAMECMT Record (0x0894)
     * Description: Defines a comment associated with a specified name.
     * REFERENCE:
     *
     * @author Andrew Shirley (aks at corefiling.co.uk)
     */
    public class NameCommentRecord : StandardRecord
    {
        public const short sid = 0x0894;

        private short field_1_record_type;
        private short field_2_frt_cell_ref_flag;
        private long field_3_reserved;
        //private short             field_4_name_Length;
        //private short             field_5_comment_Length;
        private String field_6_name_text;
        private String field_7_comment_text;

        public NameCommentRecord(string name, String comment)
        {
            field_1_record_type = 0;
            field_2_frt_cell_ref_flag = 0;
            field_3_reserved = 0;
            field_6_name_text = name;
            field_7_comment_text = comment;
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            int field_4_name_length = field_6_name_text.Length;
            int field_5_comment_length = field_7_comment_text.Length;

            out1.WriteShort(field_1_record_type);
            out1.WriteShort(field_2_frt_cell_ref_flag);
            out1.WriteLong(field_3_reserved);
            out1.WriteShort(field_4_name_length);
            out1.WriteShort(field_5_comment_length);

            out1.WriteByte(0);
            StringUtil.PutCompressedUnicode(field_6_name_text,out1);
            out1.WriteByte(0);
            StringUtil.PutCompressedUnicode(field_7_comment_text, out1);
        }

        protected override int DataSize
        {
            get
            {
                return 18 // 4 shorts + 1 long + 2 spurious 'nul's
                     + field_6_name_text.Length
                     + field_7_comment_text.Length;
            }
        }

        /**
         * @param ris the RecordInputstream to read the record from
         */
        public NameCommentRecord(RecordInputStream ris)
        {
            ILittleEndianInput in1 = ris;
            field_1_record_type = in1.ReadShort();
            field_2_frt_cell_ref_flag = in1.ReadShort();
            field_3_reserved = in1.ReadLong();
            int field_4_name_length = in1.ReadShort();
            int field_5_comment_length = in1.ReadShort();

            in1.ReadByte(); //spurious NUL
            field_6_name_text = StringUtil.ReadCompressedUnicode(in1, field_4_name_length);
            in1.ReadByte(); //spurious NUL
            field_7_comment_text = StringUtil.ReadCompressedUnicode(in1, field_5_comment_length);
        }

        /**
         * return the non static version of the id for this record.
         */

        public override short Sid
        {
            get
            {
                return sid;
            }
        }


        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("[NAMECMT]\n");
            sb.Append("    .record type            = ").Append(HexDump.ShortToHex(field_1_record_type)).Append("\n");
            sb.Append("    .frt cell ref flag      = ").Append(HexDump.ByteToHex(field_2_frt_cell_ref_flag)).Append("\n");
            sb.Append("    .reserved               = ").Append(field_3_reserved).Append("\n");
            sb.Append("    .name length            = ").Append(field_6_name_text.Length).Append("\n");
            sb.Append("    .comment length         = ").Append(field_7_comment_text.Length).Append("\n");
            sb.Append("    .name                   = ").Append(field_6_name_text).Append("\n");
            sb.Append("    .comment                = ").Append(field_7_comment_text).Append("\n");
            sb.Append("[/NAMECMT]\n");

            return sb.ToString();
        }

        /**
         * @return the name of the NameRecord to which this comment applies.
         */
        public String NameText
        {
            get
            {
                return field_6_name_text;
            }
            set 
            {
                field_6_name_text = value;
            }
        }
        /**
         * @return the text of the comment.
         */
        public String CommentText
        {
            get
            {
                return field_7_comment_text;
            }
            set 
            {
                field_7_comment_text = value;
            }
        }

        public short RecordType
        {
            get
            {
                return field_1_record_type;
            }
        }

    }
}






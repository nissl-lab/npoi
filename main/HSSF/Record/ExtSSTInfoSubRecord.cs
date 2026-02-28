
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
        

/*
 * ExtSSTInfoSubRecord.java
 *
 * Created on September 8, 2001, 8:37 PM
 */
namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * Extended SST table info subrecord
     * Contains the elements of "info" in the SST's array field
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     * @see org.apache.poi.hssf.record.ExtSSTRecord
     */

    public class InfoSubRecord
    {
        public const int ENCODED_SIZE = 8;
        public const short sid =
            0xFFF;                                             // only here for conformance, doesn't really have an sid
        private int field_1_stream_pos;          // stream pointer to the SST record
        private int field_2_bucket_sst_offset;   // don't really Understand this yet.
        private short field_3_zero;                // must be 0;

        /** Creates new ExtSSTInfoSubRecord */


        public InfoSubRecord(int streamPos, int bucketSstOffset)
        {
            field_1_stream_pos = streamPos;
            field_2_bucket_sst_offset = bucketSstOffset;
        }

        public InfoSubRecord(RecordInputStream in1)
        {
            field_1_stream_pos = in1.ReadInt();
            field_2_bucket_sst_offset = in1.ReadShort();
            field_3_zero = in1.ReadShort();
        }

        public int StreamPos
        {
            set { field_1_stream_pos = value; }
            get { return field_1_stream_pos; }
        }

        public short BucketRecordOffset
        {
            set { field_2_bucket_sst_offset = value; }
        }

        public int BucketSSTOffset
        {
            get { return field_2_bucket_sst_offset; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[EXTSST]\n");
            buffer.Append("    .streampos      = ")
                .Append(StringUtil.ToHexString(StreamPos)).Append("\n");
            buffer.Append("    .bucketsstoffset= ")
                .Append(StringUtil.ToHexString(BucketSSTOffset)).Append("\n");
            buffer.Append("    .zero           = ")
                .Append(StringUtil.ToHexString(field_3_zero)).Append("\n");
            buffer.Append("[/EXTSST]\n");
            return buffer.ToString();
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_stream_pos);
            out1.WriteShort(field_2_bucket_sst_offset);
            out1.WriteShort(field_3_zero);
        }

    }
}
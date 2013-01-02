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

    using NPOI.HSSF.Record.Common;
    using NPOI.Util;
    using System.Text;

    /**
     * Title: FeatHdr (Feature Header) Record
     * 
     * This record specifies common information for Shared Features, and 
     *  specifies the beginning of a collection of records to define them. 
     * The collection of data (Globals Substream ABNF, macro sheet substream 
     *  ABNF or worksheet substream ABNF) specifies Shared Feature data.
     */
    public class FeatHdrRecord : StandardRecord
    {
        /**
         * Specifies the enhanced protection type. Used to protect a 
         * shared workbook by restricting access to some areas of it 
         */
        public const int SHAREDFEATURES_ISFPROTECTION = 0x02;
        /**
         * Specifies that formula errors should be ignored 
         */
        public const int SHAREDFEATURES_ISFFEC2 = 0x03;
        /**
         * Specifies the smart tag type. Recognises certain
         * types of entries (proper names, dates/times etc) and
         * flags them for action 
         */
        public const int SHAREDFEATURES_ISFFACTOID = 0x04;
        /**
         * Specifies the shared list type. Used for a table
         * within a sheet
         */
        public const int SHAREDFEATURES_ISFLIST = 0x05;


        public const short sid = 0x0867;

        private FtrHeader futureHeader;
        private int isf_sharedFeatureType; // See SHAREDFEATURES_
        private byte reserved; // Should always be one
        /** 
         * 0x00000000 = rgbHdrData not present
         * 0xffffffff = rgbHdrData present
         */
        private long cbHdrData;
        /** We need a BOFRecord to make sense of this... */
        private byte[] rgbHdrData;

        public FeatHdrRecord()
        {
            futureHeader = new FtrHeader();
            futureHeader.RecordType = (/*setter*/sid);
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public FeatHdrRecord(RecordInputStream in1)
        {
            futureHeader = new FtrHeader(in1);

            isf_sharedFeatureType = in1.ReadShort();
            reserved = (byte)in1.ReadByte();
            cbHdrData = in1.ReadInt();
            // Don't process this just yet, need the BOFRecord
            rgbHdrData = in1.ReadRemainder();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[FEATURE HEADER]\n");

            // TODO ...

            buffer.Append("[/FEATURE HEADER]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            futureHeader.Serialize(out1);

            out1.WriteShort(isf_sharedFeatureType);
            out1.WriteByte(reserved);
            out1.WriteInt((int)cbHdrData);
            out1.Write(rgbHdrData);
        }

        protected override int DataSize
        {
            get
            {
                return 12 + 2 + 1 + 4 + rgbHdrData.Length;
            }
        }

        //HACK: do a "cheat" Clone, see Record.java for more information
        public override Object Clone()
        {
            return CloneViaReserialise();
        }


    }
}


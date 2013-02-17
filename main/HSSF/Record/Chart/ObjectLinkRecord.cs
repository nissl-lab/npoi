
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


namespace NPOI.HSSF.Record.Chart
{
    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * Links text to an object on the chart or identifies it as the title.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    public class ObjectLinkRecord
       : StandardRecord
    {
        public const short sid = 0x1027;
        private short field_1_anchorId;
        public const short ANCHOR_ID_CHART_TITLE = 1;
        public const short ANCHOR_ID_Y_AXIS = 2;
        public const short ANCHOR_ID_X_AXIS = 3;
        public const short ANCHOR_ID_SERIES_OR_POINT = 4;
        public const short ANCHOR_ID_Z_AXIS = 7;
        private short field_2_link1;
        private short field_3_link2;


        public ObjectLinkRecord()
        {

        }

        /**
         * Constructs a ObjectLink record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public ObjectLinkRecord(RecordInputStream in1)
        {

            field_1_anchorId = in1.ReadShort();
            field_2_link1 = in1.ReadShort();
            field_3_link2 = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[OBJECTLINK]\n");
            buffer.Append("    .AnchorId             = ")
                .Append("0x").Append(HexDump.ToHex(AnchorId))
                .Append(" (").Append(AnchorId).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .link1                = ")
                .Append("0x").Append(HexDump.ToHex(Link1))
                .Append(" (").Append(Link1).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .link2                = ")
                .Append("0x").Append(HexDump.ToHex(Link2))
                .Append(" (").Append(Link2).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/OBJECTLINK]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_anchorId);
            out1.WriteShort(field_2_link1);
            out1.WriteShort(field_3_link2);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            ObjectLinkRecord rec = new ObjectLinkRecord();

            rec.field_1_anchorId = field_1_anchorId;
            rec.field_2_link1 = field_2_link1;
            rec.field_3_link2 = field_3_link2;
            return rec;
        }




        /**
         * Get the anchor id field for the ObjectLink record.
         *
         * @return  One of 
         *        ANCHOR_ID_CHART_TITLE
         *        ANCHOR_ID_Y_AXIS
         *        ANCHOR_ID_X_AXIS
         *        ANCHOR_ID_SERIES_OR_POINT
         *        ANCHOR_ID_Z_AXIS
         */
        public short AnchorId
        {
            get
            {
                return field_1_anchorId;
            }
            set {
                this.field_1_anchorId =value;
            }
        }

        /**
         * Get the link 1 field for the ObjectLink record.
         */
        public short Link1
        {
            get
            {
                return field_2_link1;
            }
            set 
            {
                field_2_link1 = value;
            }
        }


        /**
         * Get the link 2 field for the ObjectLink record.
         */
        public short Link2
        {
            get
            {
                return field_3_link2;
            }

            set 
            {
                field_3_link2 = value;
            }
        }
    }
}

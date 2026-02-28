
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
     * The plot growth record specifies the scaling factors used when a font is scaled.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class PlotGrowthRecord: StandardRecord
    {
        public const short sid = 0x1064;
        private int field_1_horizontalScale;
        private int field_2_verticalScale;


        public PlotGrowthRecord()
        {

        }

        /**
         * Constructs a PlotGrowth record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public PlotGrowthRecord(RecordInputStream in1)
        {
            field_1_horizontalScale = in1.ReadInt();
            field_2_verticalScale = in1.ReadInt();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PLOTGROWTH]\n");
            buffer.Append("    .horizontalScale      = ")
                .Append("0x").Append(HexDump.ToHex(HorizontalScale))
                .Append(" (").Append(HorizontalScale).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .verticalScale        = ")
                .Append("0x").Append(HexDump.ToHex(VerticalScale))
                .Append(" (").Append(VerticalScale).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/PLOTGROWTH]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_horizontalScale);
            out1.WriteInt(field_2_verticalScale);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 4 + 4; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            PlotGrowthRecord rec = new PlotGrowthRecord();

            rec.field_1_horizontalScale = field_1_horizontalScale;
            rec.field_2_verticalScale = field_2_verticalScale;
            return rec;
        }

        /**
         * Get the horizontalScale field for the PlotGrowth record.
         */
        public int HorizontalScale
        {
            get
            {
                return field_1_horizontalScale;
            }
            set
            {
                this.field_1_horizontalScale =value;
            }
        }
        /**
         * Get the verticalScale field for the PlotGrowth record.
         */
        public int VerticalScale
        {
            get
            {
                return field_2_verticalScale;
            }
            set 
            {
                this.field_2_verticalScale = value;
            }
        }

    }
}


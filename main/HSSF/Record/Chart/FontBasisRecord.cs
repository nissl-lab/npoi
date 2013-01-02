
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
     * The font basis record stores various font metrics.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class FontBasisRecord
       : StandardRecord
    {
        public const short sid = 0x1060;
        private short field_1_xBasis;
        private short field_2_yBasis;
        private short field_3_heightBasis;
        private short field_4_scale;
        private short field_5_indexToFontTable;


        public FontBasisRecord()
        {

        }

        /**
         * Constructs a FontBasis record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public FontBasisRecord(RecordInputStream in1)
        {

            field_1_xBasis = in1.ReadShort();
            field_2_yBasis = in1.ReadShort();
            field_3_heightBasis = in1.ReadShort();
            field_4_scale = in1.ReadShort();
            field_5_indexToFontTable = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FBI]\n");
            buffer.Append("    .xBasis               = ")
                .Append("0x").Append(HexDump.ToHex(XBasis))
                .Append(" (").Append(XBasis).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .yBasis               = ")
                .Append("0x").Append(HexDump.ToHex(YBasis))
                .Append(" (").Append(YBasis).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .heightBasis          = ")
                .Append("0x").Append(HexDump.ToHex(HeightBasis))
                .Append(" (").Append(HeightBasis).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .scale                = ")
                .Append("0x").Append(HexDump.ToHex(Scale))
                .Append(" (").Append(Scale).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .indexToFontTable     = ")
                .Append("0x").Append(HexDump.ToHex(IndexToFontTable))
                .Append(" (").Append(IndexToFontTable).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/FBI]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_xBasis);
            out1.WriteShort(field_2_yBasis);
            out1.WriteShort(field_3_heightBasis);
            out1.WriteShort(field_4_scale);
            out1.WriteShort(field_5_indexToFontTable);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return  2 + 2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            FontBasisRecord rec = new FontBasisRecord();

            rec.field_1_xBasis = field_1_xBasis;
            rec.field_2_yBasis = field_2_yBasis;
            rec.field_3_heightBasis = field_3_heightBasis;
            rec.field_4_scale = field_4_scale;
            rec.field_5_indexToFontTable = field_5_indexToFontTable;
            return rec;
        }




        /**
         * Get the x Basis field for the FontBasis record.
         */
        public short XBasis
        {
            get
            {
                return field_1_xBasis;
            }
            set 
            {
                field_1_xBasis = value;
            }
        }

        /**
         * Get the y Basis field for the FontBasis record.
         */
        public short YBasis
        {
            get
            {
                return field_2_yBasis;
            }
            set 
            {
                field_2_yBasis = value;
            }
        }

        /**
         * Get the height basis field for the FontBasis record.
         */
        public short HeightBasis
        {
            get
            {
                return field_3_heightBasis;
            }
            set 
            {
                this.field_3_heightBasis = value;
            }
        }

        /**
         * Get the scale field for the FontBasis record.
         */
        public short Scale
        {
            get
            {
                return field_4_scale;
            }
            set 
            {
                field_4_scale = value;
            }
        }

        /**
         * Get the index to font table field for the FontBasis record.
         */
        public short IndexToFontTable
        {
            get
            {
                return field_5_indexToFontTable;
            }
            set 
            {
                this.field_5_indexToFontTable = value;
            }
        }


    }

}

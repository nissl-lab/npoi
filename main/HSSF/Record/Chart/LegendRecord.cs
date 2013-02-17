
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
     * Defines a legend for a chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    public class LegendRecord
       : StandardRecord
    {
        public const short sid = 0x1015;
        private int field_1_xAxisUpperLeft;
        private int field_2_yAxisUpperLeft;
        private int field_3_xSize;
        private int field_4_ySize;
        private byte field_5_type;
        public const byte TYPE_BOTTOM = 0;
        public const byte TYPE_CORNER = 1;
        public const byte TYPE_TOP = 2;
        public const byte TYPE_RIGHT = 3;
        public const byte TYPE_LEFT = 4;
        public const byte TYPE_UNDOCKED = 7;
        private byte field_6_spacing;
        public const byte SPACING_CLOSE = 0;
        public const byte SPACING_MEDIUM = 1;
        public const byte SPACING_OPEN = 2;
        private short field_7_options;
        private BitField autoPosition = BitFieldFactory.GetInstance(0x1);
        private BitField autoSeries = BitFieldFactory.GetInstance(0x2);
        private BitField autoXPositioning = BitFieldFactory.GetInstance(0x4);
        private BitField autoYPositioning = BitFieldFactory.GetInstance(0x8);
        private BitField vertical = BitFieldFactory.GetInstance(0x10);
        private BitField dataTable = BitFieldFactory.GetInstance(0x20);


        public LegendRecord()
        {

        }

        /**
         * Constructs a Legend record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public LegendRecord(RecordInputStream in1)
        {
            field_1_xAxisUpperLeft = in1.ReadInt();
            field_2_yAxisUpperLeft = in1.ReadInt();
            field_3_xSize = in1.ReadInt();
            field_4_ySize = in1.ReadInt();
            field_5_type = (byte)in1.ReadByte();
            field_6_spacing = (byte)in1.ReadByte();
            field_7_options = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[LEGEND]\n");
            buffer.Append("    .xAxisUpperLeft       = ")
                .Append("0x").Append(HexDump.ToHex(XAxisUpperLeft))
                .Append(" (").Append(XAxisUpperLeft).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .yAxisUpperLeft       = ")
                .Append("0x").Append(HexDump.ToHex(YAxisUpperLeft))
                .Append(" (").Append(YAxisUpperLeft).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .xSize                = ")
                .Append("0x").Append(HexDump.ToHex(XSize))
                .Append(" (").Append(XSize).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .ySize                = ")
                .Append("0x").Append(HexDump.ToHex(YSize))
                .Append(" (").Append(YSize).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .type                 = ")
                .Append("0x").Append(HexDump.ToHex(Type))
                .Append(" (").Append(Type).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .spacing              = ")
                .Append("0x").Append(HexDump.ToHex(Spacing))
                .Append(" (").Append(Spacing).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .autoPosition             = ").Append(IsAutoPosition).Append('\n');
            buffer.Append("         .autoSeries               = ").Append(IsAutoSeries).Append('\n');
            buffer.Append("         .autoXPositioning         = ").Append(IsAutoXPositioning).Append('\n');
            buffer.Append("         .autoYPositioning         = ").Append(IsAutoYPositioning).Append('\n');
            buffer.Append("         .vertical                 = ").Append(IsVertical).Append('\n');
            buffer.Append("         .dataTable                = ").Append(IsDataTable).Append('\n');

            buffer.Append("[/LEGEND]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_xAxisUpperLeft);
            out1.WriteInt(field_2_yAxisUpperLeft);
            out1.WriteInt(field_3_xSize);
            out1.WriteInt(field_4_ySize);
            out1.WriteByte(field_5_type);
            out1.WriteByte(field_6_spacing);
            out1.WriteShort(field_7_options);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 4 + 4 + 4 + 4 + 1 + 1 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            LegendRecord rec = new LegendRecord();

            rec.field_1_xAxisUpperLeft = field_1_xAxisUpperLeft;
            rec.field_2_yAxisUpperLeft = field_2_yAxisUpperLeft;
            rec.field_3_xSize = field_3_xSize;
            rec.field_4_ySize = field_4_ySize;
            rec.field_5_type = field_5_type;
            rec.field_6_spacing = field_6_spacing;
            rec.field_7_options = field_7_options;
            return rec;
        }




        /**
         * Get the x axis upper left field for the Legend record.
         */
        public int XAxisUpperLeft
        {
            get
            {
                return field_1_xAxisUpperLeft;
            }
            set 
            {
                field_1_xAxisUpperLeft = value;
            }
        }

        /**
         * Get the y axis upper left field for the Legend record.
         */
        public int YAxisUpperLeft
        {
            get
            {
                return field_2_yAxisUpperLeft;
            }
            set 
            {
                field_2_yAxisUpperLeft = value;
            }
        }

        /**
         * Get the x size field for the Legend record.
         */
        public int XSize
        {
            get { return field_3_xSize; }
            set { this.field_3_xSize = value; }
        }

        /**
         * Get the y size field for the Legend record.
         */
        public int YSize
        {
            get { return field_4_ySize; }
            set { this.field_4_ySize = value; }
        }
        /**
         * Get the type field for the Legend record.
         *
         * @return  One of 
         *        TYPE_BOTTOM
         *        TYPE_CORNER
         *        TYPE_TOP
         *        TYPE_RIGHT
         *        TYPE_LEFT
         *        TYPE_UNDOCKED
         */
        public byte Type
        {
            get { return field_5_type; }
            set { this.field_5_type = value; }
        }

        /**
         * Get the spacing field for the Legend record.
         *
         * @return  One of 
         *        SPACING_CLOSE
         *        SPACING_MEDIUM
         *        SPACING_OPEN
         */
        public byte Spacing
        {
            get { return field_6_spacing; }
            set { this.field_6_spacing = value; }
        }


        /**
         * Get the options field for the Legend record.
         */
        public short Options
        {
            get { return field_7_options; }
            set { this.field_7_options = value; }
        }

        /**
         * automatic positioning (1=docked)
         * @return  the auto position field value.
         */
        public bool IsAutoPosition
        {
            get { return autoPosition.IsSet(field_7_options); }
            set { field_7_options = autoPosition.SetShortBoolean(field_7_options, value); }
        }

        /**
         * excel 5 only (true)
         * @return  the auto series field value.
         */
        public bool IsAutoSeries
        {
            get { return autoSeries.IsSet(field_7_options); }
            set { field_7_options = autoSeries.SetShortBoolean(field_7_options, value); }
        }


        /**
         * position of legend on the x axis is automatic
         * @return  the auto x positioning field value.
         */
        public bool IsAutoXPositioning
        {
            get{return autoXPositioning.IsSet(field_7_options);}
            set { field_7_options = autoXPositioning.SetShortBoolean(field_7_options, value); }
        }

        /**
         * position of legend on the y axis is automatic
         * @return  the auto y positioning field value.
         */
        public bool IsAutoYPositioning
        {
            get{return autoYPositioning.IsSet(field_7_options);}
            set{field_7_options = autoYPositioning.SetShortBoolean(field_7_options, value);}
        }

        /**
         * vertical or horizontal legend (1 or 0 respectively).  Always 0 if not automatic.
         * @return  the vertical field value.
         */
        public bool IsVertical
        {
            get { return vertical.IsSet(field_7_options); }
            set { field_7_options = vertical.SetShortBoolean(field_7_options, value); }
        }


        /**
         * 1 if chart Contains data table
         * @return  the data table field value.
         */
        public bool IsDataTable
        {
            get{return dataTable.IsSet(field_7_options);}
            set { field_7_options = dataTable.SetShortBoolean(field_7_options, value); }
        }


    }
}
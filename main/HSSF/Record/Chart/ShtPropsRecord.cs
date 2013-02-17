
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
     * Describes a chart sheet properties record.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    /// <summary>
    /// specifies properties of a chart as defined by the Chart Sheet Substream ABNF
    /// </summary>
    public class ShtPropsRecord
       : StandardRecord
    {
        public const short sid = 0x1044;
        private short field_1_flags;
        private BitField manSerAlloc = BitFieldFactory.GetInstance(0x1);
        private BitField plotVisibleOnly = BitFieldFactory.GetInstance(0x2);
        private BitField doNotSizeWithWindow = BitFieldFactory.GetInstance(0x4);
        private BitField manPlotArea = BitFieldFactory.GetInstance(0x8);
        private BitField alwaysAutoPlotArea = BitFieldFactory.GetInstance(0x10);
        private byte field_2_mdBlank;
        private byte field_3_reserved;

        public const byte EMPTY_NOT_PLOTTED = 0;
        public const byte EMPTY_ZERO = 1;
        public const byte EMPTY_INTERPOLATED = 2;


        public ShtPropsRecord()
        {

        }

        /**
         * Constructs a SheetProperties record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public ShtPropsRecord(RecordInputStream in1)
        {

            field_1_flags = in1.ReadShort();
            field_2_mdBlank = (byte)in1.ReadByte();
            field_3_reserved = (byte)in1.ReadByte();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SHTPROPS]\n");
            buffer.Append("    .flags                = ")
                .Append("0x").Append(HexDump.ToHex(Flags))
                .Append(" (").Append(Flags).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .chartTypeManuallyFormatted     = ").Append(IsManSerAlloc).Append('\n');
            buffer.Append("         .plotVisibleOnly          = ").Append(IsPlotVisibleOnly).Append('\n');
            buffer.Append("         .doNotSizeWithWindow      = ").Append(IsNotSizeWithWindow).Append('\n');
            buffer.Append("         .defaultPlotDimensions     = ").Append(IsManPlotArea).Append('\n');
            buffer.Append("         .autoPlotArea             = ").Append(IsAlwaysAutoPlotArea).Append('\n');
            buffer.Append("    .empty                = ")
                .Append("0x").Append(HexDump.ToHex(Blank))
                .Append(" (").Append(Blank).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SHTPROPS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_flags);
            out1.WriteByte(field_2_mdBlank);
            out1.WriteByte(0);  //reserved field
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            ShtPropsRecord rec = new ShtPropsRecord();

            rec.field_1_flags = field_1_flags;
            rec.field_2_mdBlank = field_2_mdBlank;
            rec.field_3_reserved = field_3_reserved;
            return rec;
        }




        /**
         * Get the flags field for the SheetProperties record.
         */
        public short Flags
        {
            get { return field_1_flags; }
            set { this.field_1_flags = value; }
        }


        /**
         * Get the empty field for the SheetProperties record.
         *
         * @return  One of 
         *        EMPTY_NOT_PLOTTED
         *        EMPTY_ZERO
         *        EMPTY_INTERPOLATED
         */
        /// <summary>
        /// specifies how the empty cells are plotted be a value from the following table:
        /// 0x00   Empty cells are not plotted.
        /// 0x01   Empty cells are plotted as zero.
        /// 0x02   Empty cells are plotted as interpolated.
        /// </summary>
        public byte Blank
        {
            get { return field_2_mdBlank; }
            set { this.field_2_mdBlank = value; }
        }


        /// <summary>
        /// whether series are automatically allocated for the chart.
        /// </summary>
        public bool IsManSerAlloc
        {
            get { return manSerAlloc.IsSet(field_1_flags); }
            set { field_1_flags = manSerAlloc.SetShortBoolean(field_1_flags, value); }
        }

        /// <summary>
        /// whether to plot visible cells only.
        /// </summary>
        public bool IsPlotVisibleOnly
        {
            get { return plotVisibleOnly.IsSet(field_1_flags); }
            set { field_1_flags = plotVisibleOnly.SetShortBoolean(field_1_flags, value); }
        }

        /// <summary>
        /// whether to size the chart with the window.
        /// </summary>
        public bool IsNotSizeWithWindow
        {
            get { return doNotSizeWithWindow.IsSet(field_1_flags); }
            set { field_1_flags = doNotSizeWithWindow.SetShortBoolean(field_1_flags, value); }
        }

        /// <summary>
        /// If fAlwaysAutoPlotArea is 1, then this field MUST be 1. 
        /// If fAlwaysAutoPlotArea is 0, then this field MUST be ignored.
        /// </summary>
        public bool IsManPlotArea
        {
            get
            {
                return manPlotArea.IsSet(field_1_flags);
            }
            set { field_1_flags = manPlotArea.SetShortBoolean(field_1_flags, value); }
        }

        /// <summary>
        /// specifies whether the default plot area dimension (2) is used.
        /// 0  Use the default plot area dimension (2) regardless of the Pos record information.
        /// 1  Use the plot area dimension (2) of the Pos record; and fManPlotArea MUST be 1.
        /// </summary>
        public bool IsAlwaysAutoPlotArea
        {
            get
            {
                return alwaysAutoPlotArea.IsSet(field_1_flags);
            }
            set
            {
                field_1_flags = alwaysAutoPlotArea.SetShortBoolean(field_1_flags, value);
                if (value)
                    IsManPlotArea = value;
            }
        }


    }
}

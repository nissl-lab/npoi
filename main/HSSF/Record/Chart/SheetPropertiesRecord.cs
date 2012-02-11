
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
    public class SheetPropertiesRecord
       : StandardRecord
    {
        public static short sid = 0x1044;
        private short field_1_flags;
        private BitField chartTypeManuallyFormatted = BitFieldFactory.GetInstance(0x1);
        private BitField plotVisibleOnly = BitFieldFactory.GetInstance(0x2);
        private BitField doNotSizeWithWindow = BitFieldFactory.GetInstance(0x4);
        private BitField defaultPlotDimensions = BitFieldFactory.GetInstance(0x8);
        private BitField autoPlotArea = BitFieldFactory.GetInstance(0x10);
        private byte field_2_empty;
        private byte field_3_reserved;

        public static byte EMPTY_NOT_PLOTTED = 0;
        public static byte EMPTY_ZERO = 1;
        public static byte EMPTY_INTERPOLATED = 2;


        public SheetPropertiesRecord()
        {

        }

        /**
         * Constructs a SheetProperties record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SheetPropertiesRecord(RecordInputStream in1)
        {

            field_1_flags = in1.ReadShort();
            field_2_empty = (byte)in1.ReadByte();
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
            buffer.Append("         .chartTypeManuallyFormatted     = ").Append(IsChartTypeManuallyFormatted).Append('\n');
            buffer.Append("         .plotVisibleOnly          = ").Append(IsPlotVisibleOnly).Append('\n');
            buffer.Append("         .doNotSizeWithWindow      = ").Append(IsDoNotSizeWithWindow).Append('\n');
            buffer.Append("         .defaultPlotDimensions     = ").Append(IsDefaultPlotDimensions).Append('\n');
            buffer.Append("         .autoPlotArea             = ").Append(IsAutoPlotArea).Append('\n');
            buffer.Append("    .empty                = ")
                .Append("0x").Append(HexDump.ToHex(Empty))
                .Append(" (").Append(Empty).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SHTPROPS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_flags);
            out1.WriteByte(field_2_empty);
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
            SheetPropertiesRecord rec = new SheetPropertiesRecord();

            rec.field_1_flags = field_1_flags;
            rec.field_2_empty = field_2_empty;
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
        public byte Empty
        {
            get { return field_2_empty; }
            set { this.field_2_empty = value; }
        }


        /**
         * Has the chart type been manually formatted?
         * @return  the chart type manually formatted field value.
         */
        public bool IsChartTypeManuallyFormatted
        {
            get { return chartTypeManuallyFormatted.IsSet(field_1_flags); }
            set { field_1_flags = chartTypeManuallyFormatted.SetShortBoolean(field_1_flags, value); }
        }

        /**
         * Only show visible cells on the chart.
         * @return  the plot visible only field value.
         */
        public bool IsPlotVisibleOnly
        {
            get { return plotVisibleOnly.IsSet(field_1_flags); }
            set { field_1_flags = plotVisibleOnly.SetShortBoolean(field_1_flags, value); }
        }

        /**
         * Do not size the chart when the window Changes size
         * @return  the do not size with window field value.
         */
        public bool IsDoNotSizeWithWindow
        {
            get { return doNotSizeWithWindow.IsSet(field_1_flags); }
            set { field_1_flags = doNotSizeWithWindow.SetShortBoolean(field_1_flags, value); }
        }

        /**
         * Indicates that the default area dimensions should be used.
         * @return  the default plot dimensions field value.
         */
        public bool IsDefaultPlotDimensions
        {
            get
            {
                return defaultPlotDimensions.IsSet(field_1_flags);
            }
            set { field_1_flags = defaultPlotDimensions.SetShortBoolean(field_1_flags, value); }
        }

        /**
         * ??
         * @return  the auto plot area field value.
         */
        public bool IsAutoPlotArea
        {
            get
            {
                return autoPlotArea.IsSet(field_1_flags);
            }
            set
            { field_1_flags = autoPlotArea.SetShortBoolean(field_1_flags, value); }
        }


    }
}

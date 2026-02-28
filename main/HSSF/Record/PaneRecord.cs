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

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * Describes the frozen and Unfozen panes.
     * NOTE: This source Is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class PaneRecord : StandardRecord
    {
        public const short sid = 0x41;
        private short field_1_x;
        private short field_2_y;
        private short field_3_topRow;
        private short field_4_leftColumn;
        private short field_5_activePane;
        public const short ACTIVE_PANE_LOWER_RIGHT = 0;
        public const short ACTIVE_PANE_UPPER_RIGHT = 1;
        public const short ACTIVE_PANE_LOWER_LEFT = 2;
        public const short ACTIVE_PANE_UPPER_LEFT = 3;


        public PaneRecord()
        {

        }

        /**
         * Constructs a Pane record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public PaneRecord(RecordInputStream in1)
        {
            field_1_x = in1.ReadShort();
            field_2_y = in1.ReadShort();
            field_3_topRow = in1.ReadShort();
            field_4_leftColumn = in1.ReadShort();
            field_5_activePane = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PANE]\n");
            buffer.Append("    .x                    = ")
                .Append("0x").Append(HexDump.ToHex(X))
                .Append(" (").Append(X).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .y                    = ")
                .Append("0x").Append(HexDump.ToHex(Y))
                .Append(" (").Append(Y).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .topRow               = ")
                .Append("0x").Append(HexDump.ToHex(TopRow))
                .Append(" (").Append(TopRow).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .leftColumn           = ")
                .Append("0x").Append(HexDump.ToHex(LeftColumn))
                .Append(" (").Append(LeftColumn).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .activePane           = ")
                .Append("0x").Append(HexDump.ToHex(ActivePane))
                .Append(" (").Append(ActivePane).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/PANE]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_x);
            out1.WriteShort(field_2_y);
            out1.WriteShort(field_3_topRow);
            out1.WriteShort(field_4_leftColumn);
            out1.WriteShort(field_5_activePane);
        }

        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2 + 2 + 2;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            PaneRecord rec = new PaneRecord();

            rec.field_1_x = field_1_x;
            rec.field_2_y = field_2_y;
            rec.field_3_topRow = field_3_topRow;
            rec.field_4_leftColumn = field_4_leftColumn;
            rec.field_5_activePane = field_5_activePane;
            return rec;
        }




        /**
         * Get the x field for the Pane record.
         */
        public short X
        {
            get
            {
                return field_1_x;
            }
            set 
            {
                this.field_1_x = value;
            }
        }

        /**
         * Get the y field for the Pane record.
         */
        public short Y
        {
            get
            {
                return field_2_y;
            }
            set 
            {
                this.field_2_y = value;
            }
        }


        /**
         * Get the top row field for the Pane record.
         */
        public short TopRow
        {
            get
            {
                return field_3_topRow;
            }
            set 
            {
                this.field_3_topRow = value;
            }
        }
        /**
         * Get the left column field for the Pane record.
         */
        public short LeftColumn
        {
            get
            {
                return field_4_leftColumn;
            }
            set 
            {
                this.field_4_leftColumn = value;
            }
        }

        /**
         * Get the active pane field for the Pane record.
         *
         * @return  One of 
         *        ACTIVE_PANE_LOWER_RIGHT
         *        ACTIVE_PANE_UPPER_RIGHT
         *        ACTIVE_PANE_LOWER_LEFT
         *        ACTIVE_PANE_UPPER_LEFT
         */
        public short ActivePane
        {
            get
            {
                return field_5_activePane;
            }
            set 
            {
                this.field_5_activePane = value;
            }
        }
    }
}
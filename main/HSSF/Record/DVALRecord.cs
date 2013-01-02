/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed Under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

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
     * Title:        DATAVALIDATIONS Record
     * Description:  used in data validation ;
     *               This record Is the list header of all data validation records (0x01BE) in the current sheet.
     * @author Dragos Buleandra (dragos.buleandra@trade2b.ro)
     */
    public class DVALRecord : StandardRecord
    {
        public const short sid = 0x01B2;

        /** Options of the DVAL */
        private short field_1_options;
        /** Horizontal position of the dialog */
        private int field_2_horiz_pos;
        /** Vertical position of the dialog */
        private int field_3_vert_pos;

        /** Object ID of the drop down arrow object for list boxes ;
         * in our case this will be always FFFF , Until
         * MSODrawingGroup and MSODrawing records are implemented */
        private int field_cbo_id;

        /** Number of following DV Records */
        private int field_5_dv_no;

        public DVALRecord()
        {
            field_cbo_id = unchecked((int)0xFFFFFFFF);
            field_5_dv_no = 0x00000000;
        }

        /**
         * Constructs a DVAL record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public DVALRecord(RecordInputStream in1)
        {
            this.field_1_options = in1.ReadShort();
            this.field_2_horiz_pos = in1.ReadInt();
            this.field_3_vert_pos = in1.ReadInt();
            this.field_cbo_id = in1.ReadInt();
            this.field_5_dv_no = in1.ReadInt();
        }


        /**
         * @return the field_1_options
         */
        public short Options
        {
            get { return field_1_options; }
            set { this.field_1_options = value; }
        }

        /**
         * @return the Horizontal position of the dialog
         */
        public int HorizontalPos
        {
            get
            {
                return field_2_horiz_pos;
            }
            set 
            {
                this.field_2_horiz_pos = value;
            }
        }

        /**
         * @return the the Vertical position of the dialog
         */
        public int VerticalPos
        {
            get
            {
                return field_3_vert_pos;
            }
            set 
            {
                this.field_3_vert_pos = value;
            }
        }

        /**
         * Get Object ID of the drop down arrow object for list boxes
         */
        public int ObjectID
        {
            get
            {
                return this.field_cbo_id;
            }
            set 
            {
                this.field_cbo_id = value;
            }
        }

        /**
         * Get number of following DV records
         */
        public int DVRecNo
        {
            get
            {
                return this.field_5_dv_no;
            }
            set 
            {
                this.field_5_dv_no = value;
            }
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DVAL]\n");
            buffer.Append("    .options      = ").Append(this.Options).Append('\n');
            buffer.Append("    .horizPos     = ").Append(this.HorizontalPos).Append('\n');
            buffer.Append("    .vertPos      = ").Append(this.VerticalPos).Append('\n');
            buffer.Append("    .comboObjectID   = ").Append(StringUtil.ToHexString(this.ObjectID)).Append("\n");
            buffer.Append("    .DVRecordsNumber = ").Append(StringUtil.ToHexString(this.DVRecNo)).Append("\n");
            buffer.Append("[/DVAL]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
		    out1.WriteShort(Options);
		    out1.WriteInt(HorizontalPos);
		    out1.WriteInt(VerticalPos);
		    out1.WriteInt(ObjectID);
		    out1.WriteInt(DVRecNo);
        }

        protected override int DataSize
        {
            get { return 18; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            DVALRecord rec = new DVALRecord();
            rec.field_1_options = field_1_options;
            rec.field_2_horiz_pos = field_2_horiz_pos;
            rec.field_3_vert_pos = field_3_vert_pos;
            rec.field_cbo_id = this.field_cbo_id;
            rec.field_5_dv_no = this.field_5_dv_no;
            return rec;
        }
    }
}
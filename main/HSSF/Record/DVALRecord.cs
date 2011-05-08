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
    public class DVALRecord : Record
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
         * @param field_1_options the options of the dialog
         */
        public void SetOptions(short field_1_options)
        {
            this.field_1_options = field_1_options;
        }

        /**
         * @param field_2_horiz_pos the Horizontal position of the dialog
         */
        public void SetHorizontalPos(int field_2_horiz_pos)
        {
            this.field_2_horiz_pos = field_2_horiz_pos;
        }

        /**
         * @param field_3_vert_pos the Vertical position of the dialog
         */
        public void SetVerticalPos(int field_3_vert_pos)
        {
            this.field_3_vert_pos = field_3_vert_pos;
        }

        /**
         * Set the object ID of the drop down arrow object for list boxes
         * @param cboID - Object ID
         */
        public void SetObjectID(int cboID)
        {
            this.field_cbo_id = cboID;
        }

        /**
         * Set the number of following DV records
         * @param dvNo - the DV records number
         */
        public void SetDVRecNo(int dvNo)
        {
            this.field_5_dv_no = dvNo;
        }



        /**
         * @return the field_1_options
         */
        public short Options
        {
            get { return field_1_options; }
        }

        /**
         * @return the Horizontal position of the dialog
         */
        public int GetHorizontalPos()
        {
            return field_2_horiz_pos;
        }

        /**
         * @return the the Vertical position of the dialog
         */
        public int GetVerticalPos()
        {
            return field_3_vert_pos;
        }

        /**
         * Get Object ID of the drop down arrow object for list boxes
         */
        public int GetObjectID()
        {
            return this.field_cbo_id;
        }

        /**
         * Get number of following DV records
         */
        public int GetDVRecNo()
        {
            return this.field_5_dv_no;
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DVAL]\n");
            buffer.Append("    .options      = ").Append(this.Options).Append('\n');
            buffer.Append("    .horizPos     = ").Append(this.GetHorizontalPos()).Append('\n');
            buffer.Append("    .vertPos      = ").Append(this.GetVerticalPos()).Append('\n');
            buffer.Append("    .comboObjectID   = ").Append(StringUtil.ToHexString(this.GetObjectID())).Append("\n");
            buffer.Append("    .DVRecordsNumber = ").Append(StringUtil.ToHexString(this.GetDVRecNo())).Append("\n");
            buffer.Append("[/DVAL]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, this.Sid);
            LittleEndian.PutShort(data, 2 + offset, (short)(this.RecordSize - 4));

            LittleEndian.PutShort(data, 4 + offset, this.Options);
            LittleEndian.PutInt(data, 6 + offset, this.GetHorizontalPos());
            LittleEndian.PutInt(data, 10 + offset, this.GetVerticalPos());
            LittleEndian.PutInt(data, 14 + offset, this.GetObjectID());
            LittleEndian.PutInt(data, 18 + offset, this.GetDVRecNo());
            return RecordSize;
        }

        //with 4 bytes header
        public override int RecordSize
        {
            get { return 22; }
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
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
    using System.Text;
    using System;
    using NPOI.Util;

    /**
     * NOTE: Comment Associated with a Cell (1Ch)
     *
     * @author Yegor Kozlov
     */
    public class NoteRecord : Record
    {
        public static NoteRecord[] EMPTY_ARRAY = { };
        public const short sid = 0x1C;

        /**
         * Flag indicating that the comment Is hidden (default)
         */
        public static short NOTE_HIDDEN = 0x0;

        /**
         * Flag indicating that the comment Is visible
         */
        public static short NOTE_VISIBLE = 0x2;

        private int field_1_row;
        private int field_2_col;
        private short field_3_flags;
        private short field_4_shapeid;
        private String field_5_author;

        /**
         * Construct a new <c>NoteRecord</c> and
         * Fill its data with the default values
         */
        public NoteRecord()
        {
            field_5_author = "";
            field_3_flags = 0;
        }

        /**
         * Constructs a <c>NoteRecord</c> and Fills its fields
         * from the supplied <c>RecordInputStream</c>.
         *
         * @param in the stream to Read from
         */
        public NoteRecord(RecordInputStream in1)
        {
            field_1_row = in1.ReadShort();
            field_2_col = in1.ReadShort();
            field_3_flags = in1.ReadShort();
            field_4_shapeid = in1.ReadShort();
            int Length = in1.ReadShort();
            byte[] bytes = in1.ReadRemainder();
            field_5_author = Encoding.UTF8.GetString(bytes, 1, Length);


        }

        /**
         * @return id of this record.
         */
        public override short Sid
        {
            get { return sid; }
        }

        /**
         * Serialize the record data into the supplied array of bytes
         *
         * @param offset offset in the <c>data</c>
         * @param data the data to Serialize into
         *
         * @return size of the record
         */
        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset, (short)(RecordSize - 4));

            LittleEndian.PutUShort(data, 4 + offset, field_1_row);
            LittleEndian.PutUShort(data, 6 + offset, field_2_col);
            LittleEndian.PutShort(data, 8 + offset, field_3_flags);
            LittleEndian.PutShort(data, 10 + offset, field_4_shapeid);
            LittleEndian.PutShort(data, 12 + offset, (short)field_5_author.Length);

            byte[] str = Encoding.UTF8.GetBytes(field_5_author);
            Array.Copy(str, 0, data, 15 + offset, str.Length);

            return RecordSize;
        }

        /**
         * Size of record
         */
        public override int RecordSize
        {
            get
            {
                int retval = 4 + 2 + 2 + 2 + 2 + 2 + 1 + field_5_author.Length + 1;

                return retval;

            }
        }

        /**
         * Convert this record to string.
         * Used by BiffViewer and other utulities.
         */
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[NOTE]\n");
            buffer.Append("    .recordid = 0x" + StringUtil.ToHexString(Sid) + ", size = " + RecordSize + "\n");
            buffer.Append("    .row =     " + field_1_row + "\n");
            buffer.Append("    .col =     " + field_2_col + "\n");
            buffer.Append("    .flags =   " + field_3_flags + "\n");
            buffer.Append("    .shapeid = " + field_4_shapeid + "\n");
            buffer.Append("    .author =  " + field_5_author + "\n");
            buffer.Append("[/NOTE]\n");
            return buffer.ToString();
        }

        /**
         * Return the row that Contains the comment
         *
         * @return the row that Contains the comment
         */
        public int Row
        {
            get{return field_1_row;}
            set{ field_1_row = value;}
        }

        /**
         * Return the column that Contains the comment
         *
         * @return the column that Contains the comment
         */
        public int Column
        {
            get { return field_2_col; }
            set { field_2_col = value; }
        }

        /**
         * Options flags.
         *
         * @return the options flag
         * @see #NOTE_VISIBLE
         * @see #NOTE_HIDDEN
         */
        public short Flags
        {
            get { return field_3_flags; }
            set { field_3_flags = value; }
        }

        /**
         * Object id for OBJ record that Contains the comment
         */
        public short ShapeId
        {
            get { return field_4_shapeid; }
            set { field_4_shapeid = value; }
        }


        /**
         * Name of the original comment author
         *
         * @return the name of the original author of the comment
         */
        public String Author
        {
            get { return field_5_author; }
            set { field_5_author = value; }
        }


        public override Object Clone()
        {
            NoteRecord rec = new NoteRecord();
            rec.field_1_row = field_1_row;
            rec.field_2_col = field_2_col;
            rec.field_3_flags = field_3_flags;
            rec.field_4_shapeid = field_4_shapeid;
            rec.field_5_author = field_5_author;
            return rec;
        }

    }
}
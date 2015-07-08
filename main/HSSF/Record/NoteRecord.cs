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
    using System.Globalization;


    /**
     * NOTE: Comment Associated with a Cell (1Ch)
     *
     * @author Yegor Kozlov
     */
    public class NoteRecord : StandardRecord
    {
        public static readonly NoteRecord[] EMPTY_ARRAY = { };
        public const short sid = 0x1C;

        /**
         * Flag indicating that the comment Is hidden (default)
         */
        public const short NOTE_HIDDEN = 0x0;

        /**
         * Flag indicating that the comment Is visible
         */
        public const short NOTE_VISIBLE = 0x2;

        private int field_1_row;
        private int field_2_col;
        private short field_3_flags;
        private int field_4_shapeid;
        private bool field_5_hasMultibyte;
        private String field_6_author;
        private const Byte DEFAULT_PADDING = (byte)0;

        /**
 * Saves padding byte value to reduce delta during round-trip serialization.<br/>
 *
 * The documentation is not clear about how padding should work.  In any case
 * Excel(2007) does something different.
 */
        private Byte? field_7_padding;
        /**
         * Construct a new <c>NoteRecord</c> and
         * Fill its data with the default values
         */
        public NoteRecord()
        {
            field_6_author = "";
            field_3_flags = 0;
            field_7_padding = DEFAULT_PADDING; // seems to be always present regardless of author text  
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
            field_2_col = in1.ReadUShort();
            field_3_flags = in1.ReadShort();
            field_4_shapeid = in1.ReadUShort();
            int length = in1.ReadShort();
		    field_5_hasMultibyte = in1.ReadByte() != 0x00;
		    if (field_5_hasMultibyte) {
			    field_6_author = StringUtil.ReadUnicodeLE(in1, length);
		    } else {
			    field_6_author = StringUtil.ReadCompressedUnicode(in1, length);
		    }
 		    if (in1.Available() == 1) {
			    field_7_padding = (byte)in1.ReadByte();
		    }
            else if (in1.Available() == 2 && length == 0)
            {
                // If there's no author, may be double padded
                field_7_padding = (byte)in1.ReadByte();
                in1.ReadByte();
            }
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
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_row);
		    out1.WriteShort(field_2_col);
		    out1.WriteShort(field_3_flags);
		    out1.WriteShort(field_4_shapeid);
		    out1.WriteShort(field_6_author.Length);
		    out1.WriteByte(field_5_hasMultibyte ? 0x01 : 0x00);
		    if (field_5_hasMultibyte) {
			    StringUtil.PutUnicodeLE(field_6_author, out1);
		    } else {
			    StringUtil.PutCompressedUnicode(field_6_author, out1);
		    }
		    if (field_7_padding != null) {
			    out1.WriteByte(Convert.ToInt32(field_7_padding, CultureInfo.InvariantCulture));
		    }

        }

        /**
         * Size of record
         */
        protected override int DataSize
        {
            get
            {
                return 11 // 5 shorts + 1 byte
                    + field_6_author.Length * (field_5_hasMultibyte ? 2 : 1)
                    + (field_7_padding == null ? 0 : 1);
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
            buffer.Append("    .author =  " + field_6_author + "\n");
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
        public int ShapeId
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
            get { return field_6_author; }
            set
            {
                field_6_author = value;
                field_5_hasMultibyte = StringUtil.HasMultibyte(value);
            }
        }

        /**
 * For unit testing only!
 */
        internal bool AuthorIsMultibyte
        {
            get
            {
                return field_5_hasMultibyte;
            }
        }

        public override Object Clone()
        {
            NoteRecord rec = new NoteRecord();
            rec.field_1_row = field_1_row;
            rec.field_2_col = field_2_col;
            rec.field_3_flags = field_3_flags;
            rec.field_4_shapeid = field_4_shapeid;
            rec.field_6_author = field_6_author;
            return rec;
        }

    }
}
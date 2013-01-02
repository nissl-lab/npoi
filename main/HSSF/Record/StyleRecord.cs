
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
     * Title:        Style Record
     * Description:  Describes a builtin to the gui or user defined style
     * REFERENCE:  PG 390 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author aviks : string fixes for UserDefined Style
     * @version 2.0-pre
     */

    public class StyleRecord
       : StandardRecord
    {
        public const short sid = 0x293;

	    private static BitField styleIndexMask = BitFieldFactory.GetInstance(0x0FFF);
	    private static BitField isBuiltinFlag  = BitFieldFactory.GetInstance(0x8000);

        // shared by both user defined and builtin styles
        private int field_1_xf_index;   // TODO: bitfield candidate

        // only for built in styles
        private int field_2_builtin_style;
        private int field_3_outline_style_level;

        // only for user defined styles
        private bool field_3_stringHasMultibyte;
        private String field_4_name;

        public StyleRecord()
        {
            field_1_xf_index = isBuiltinFlag.Set(field_1_xf_index);
        }
        public bool IsBuiltin
        {
            get
            {
                return isBuiltinFlag.IsSet(field_1_xf_index);
            }
        }
        /**
         * Constructs a Style record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public StyleRecord(RecordInputStream in1)
        {
            field_1_xf_index = in1.ReadShort();
            if (IsBuiltin)
            {
                field_2_builtin_style = in1.ReadByte();
                field_3_outline_style_level = in1.ReadByte();
            }
            else
            {
                int field_2_name_length = in1.ReadShort();

                // Some files from Crystal Reports lack
                //  the remaining fields, which Is naughty
                if (in1.Remaining <1)
                {
                    // Some files from Crystal Reports lack the is16BitUnicode byte
                    //  the remaining fields, which is naughty
                    if (field_2_name_length != 0)
                    {
                        throw new RecordFormatException("Ran out of data reading style record");
                    }
                    // guess this is OK if the string length is zero
                    field_4_name = "";
                }
                else
                {
				    field_3_stringHasMultibyte = in1.ReadByte() != 0x00;
				    if (field_3_stringHasMultibyte)
                    {
                        field_4_name = StringUtil.ReadUnicodeLE(in1, field_2_name_length);
                    }
                    else
                    {
                        field_4_name = StringUtil.ReadCompressedUnicode(in1,field_2_name_length);
                    }
                }
            }

            // todo sanity Check exception to make sure we're one or the other
        }

        /**
         * if this is a builtin style set the number of the built in style
         * @param  builtinStyleId style number (0-7)
         *
         */
        public void SetBuiltinStyle(int builtinStyleId)
        {
            field_1_xf_index = isBuiltinFlag.Set(field_1_xf_index);
            field_2_builtin_style = builtinStyleId;
        }
        // bitfields for field 1

        /**
         * Get the actual index of the style extended format record
         * @see #Index
         * @return index of the xf record
         */

        public short XFIndex
        {
            get { return (short)(field_1_xf_index & 0x1FFF); }
            set { field_1_xf_index = SetField(field_1_xf_index, value, 0x1FFF, 0); }
        }

        // end bitfields
        // only for user defined records
        /**
         * Get the style's name
         * @return name of the style
         * @see #NameLength
         */

        public String Name
        {
            get { return field_4_name; }
            set { 
                field_4_name = value;
                field_3_stringHasMultibyte = StringUtil.HasMultibyte(value);
                field_1_xf_index = isBuiltinFlag.Clear(field_1_xf_index);
            }
        }

        // end user defined
 
 
        /**
         * Get the row or column level of the style (if builtin 1||2)
         */

        public int OutlineStyleLevel
        {
            get
            {
                return field_3_outline_style_level;
            }
            set { field_3_outline_style_level = value & 0x00FF; }
        }

        // end builtin records
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[STYLE]\n");
            buffer.Append("    .xf_index_raw    = ")
                .Append(HexDump.ShortToHex(field_1_xf_index)).Append("\n");
            buffer.Append("        .type        = ")
                .Append(IsBuiltin?"built-in":"user-defined").Append("\n");
            buffer.Append("        .xf_index    = ")
                .Append(HexDump.ShortToHex(XFIndex)).Append("\n");
            if (IsBuiltin)
            {
                buffer.Append("    .builtin_style   = ").Append(HexDump.ByteToHex(field_2_builtin_style)).Append("\n");
                buffer.Append("    .outline_level   = ").Append(HexDump.ByteToHex(field_3_outline_style_level)).Append("\n");
            }
            else
            {
                buffer.Append("    .name            = ").Append(Name).Append("\n");
            }
            buffer.Append("[/STYLE]\n");
            return buffer.ToString();
        }

        private short SetField(int fieldValue, int new_value, int mask,
                               int ShiftLeft)
        {
            return (short)((fieldValue & ~mask)
                              | ((new_value << ShiftLeft) & mask));
        }

        public override void Serialize(ILittleEndianOutput o)
        {
            o.WriteShort(field_1_xf_index);
            if (IsBuiltin)
            {
                o.WriteByte(field_2_builtin_style);
                o.WriteByte(field_3_outline_style_level);
            }
            else
            {
                o.WriteShort(field_4_name.Length);
                o.WriteByte(field_3_stringHasMultibyte ? 0x01 : 0x00);
                if (field_3_stringHasMultibyte)
                {
                    StringUtil.PutUnicodeLE(Name, o);
                }
                else
                {
                    StringUtil.PutCompressedUnicode(Name, o);
                }
            }
        }

        protected override int DataSize
        {
            get
            {
                if (IsBuiltin)
                {
                    return 4; // short, byte, byte
                }
                return 2 // short xf index 
                    + 3 // str len + flag 
                    + field_4_name.Length * (field_3_stringHasMultibyte ? 2 : 1);
            }
        }


        public override short Sid
        {
            get { return sid; }
        }
    }
}
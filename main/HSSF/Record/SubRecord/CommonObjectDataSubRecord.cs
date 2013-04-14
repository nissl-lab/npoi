
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
     * The common object data record is used to store all common preferences for an excel object.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public enum CommonObjectType:short
    { 
        Group = 0,
        Line = 1,
        Rectangle = 2,
        Oval = 3,
        Arc = 4,
        Chart = 5,
        Text = 6,
        Button = 7,
        Picture = 8,
        Polygon = 9,
        Reserved1 = 10,
        Checkbox = 11,
        OptionButton = 12,
        EditBox = 13,
        Label = 14,
        DialogBox = 15,
        Spinner = 16,
        ScrollBar = 17,
        ListBox = 18,
        GroupBox = 19,
        ComboBox = 20,
        Reserved2 = 21,
        Reserved3 = 22,
        Reserved4 = 23,
        Reserved5 = 24,
        Comment = 25,
        Reserved6 = 26,
        Reserved7 = 27,
        Reserved8 = 28,
        Reserved9 = 29,
        MicrosoftOfficeDrawing = 30,
    }

    public class CommonObjectDataSubRecord
       : SubRecord
    {
        public const short sid = 0x15;
        private short field_1_objectType;

        private int field_2_objectId;
        private short field_3_option;
        private BitField locked = BitFieldFactory.GetInstance(0x1);
        private BitField printable = BitFieldFactory.GetInstance(0x10);
        private BitField autoFill = BitFieldFactory.GetInstance(0x2000);
        private BitField autoline = BitFieldFactory.GetInstance(0x4000);
        private int field_4_reserved1;
        private int field_5_reserved2;
        private int field_6_reserved3;


        public CommonObjectDataSubRecord()
        {

        }

        /**
         * Constructs a CommonObjectData record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public CommonObjectDataSubRecord(ILittleEndianInput in1, int size)
        {
            if (size != 18)
            {
                throw new RecordFormatException("Expected size 18 but got (" + size + ")");
            }
            field_1_objectType = in1.ReadShort();
            field_2_objectId = in1.ReadUShort();
            field_3_option = in1.ReadShort();
            field_4_reserved1 = in1.ReadInt();
            field_5_reserved2 = in1.ReadInt();
            field_6_reserved3 = in1.ReadInt();
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ftCmo]\n");
            buffer.Append("    .objectType           = ")
                .Append("0x").Append(HexDump.ToHex((short)ObjectType))
                .Append(" (").Append(ObjectType).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .objectId             = ")
                .Append("0x").Append(HexDump.ToHex(ObjectId))
                .Append(" (").Append(ObjectId).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .option               = ")
                .Append("0x").Append(HexDump.ToHex(Option))
                .Append(" (").Append(Option).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .locked                   = ").Append(IsLocked).Append('\n');
            buffer.Append("         .printable                = ").Append(IsPrintable).Append('\n');
            buffer.Append("         .autoFill                 = ").Append(IsAutoFill).Append('\n');
            buffer.Append("         .autoline                 = ").Append(IsAutoline).Append('\n');
            buffer.Append("    .reserved1            = ")
                .Append("0x").Append(HexDump.ToHex(Reserved1))
                .Append(" (").Append(Reserved1).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .reserved2            = ")
                .Append("0x").Append(HexDump.ToHex(Reserved2))
                .Append(" (").Append(Reserved2).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .reserved3            = ")
                .Append("0x").Append(HexDump.ToHex(Reserved3))
                .Append(" (").Append(Reserved3).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/ftCmo]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(DataSize);

            out1.WriteShort(field_1_objectType);
            out1.WriteShort(field_2_objectId);
            out1.WriteShort(field_3_option);
            out1.WriteInt(field_4_reserved1);
            out1.WriteInt(field_5_reserved2);
            out1.WriteInt(field_6_reserved3);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public override int DataSize
        {
            get { return  2 + 2 + 2 + 4 + 4 + 4; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            CommonObjectDataSubRecord rec = new CommonObjectDataSubRecord();

            rec.field_1_objectType = field_1_objectType;
            rec.field_2_objectId = field_2_objectId;
            rec.field_3_option = field_3_option;
            rec.field_4_reserved1 = field_4_reserved1;
            rec.field_5_reserved2 = field_5_reserved2;
            rec.field_6_reserved3 = field_6_reserved3;
            return rec;
        }


        /**
         * Get the object type field for the CommonObjectData record.
         */
        public CommonObjectType ObjectType
        {
            get
            {
                return (CommonObjectType)field_1_objectType;
            }
            set { this.field_1_objectType = (short)value; }
        }
        /**
         * Get the object id field for the CommonObjectData record.
         */
        public int ObjectId
        {
            get
            {
                return field_2_objectId;
            }
            set { this.field_2_objectId = value; }
        }

        /**
         * Get the option field for the CommonObjectData record.
         */
        public short Option
        {
            get
            {
                return field_3_option;
            }
            set { this.field_3_option = value; }
        }

        /**
         * Get the reserved1 field for the CommonObjectData record.
         */
        public int Reserved1
        {
            get
            {
                return field_4_reserved1;
            }
            set { this.field_4_reserved1 = value; }
        }

        /**
         * Get the reserved2 field for the CommonObjectData record.
         */
        public int Reserved2
        {
            get
            {
                return field_5_reserved2;
            }
            set { this.field_5_reserved2 = value; }
        }

        /**
         * Get the reserved3 field for the CommonObjectData record.
         */
        public int Reserved3
        {
            get
            {
                return field_6_reserved3;
            }
            set { this.field_6_reserved3 = value; }
        }

        /**
         * true if object is locked when sheet has been protected
         * @return  the locked field value.
         */
        public bool IsLocked
        {
            get
            {
                return locked.IsSet(field_3_option);
            }
            set { field_3_option = locked.SetShortBoolean(field_3_option, value); }
        }

        /**
         * object appears when printed
         * @return  the printable field value.
         */
        public bool IsPrintable
        {
            get
            {
                return printable.IsSet(field_3_option);
            }
            set { field_3_option = printable.SetShortBoolean(field_3_option, value); }
        }

        /**
         * whether object uses an automatic Fill style
         * @return  the autoFill field value.
         */
        public bool IsAutoFill
        {
            get
            {
                return autoFill.IsSet(field_3_option);
            }
            set { field_3_option = autoFill.SetShortBoolean(field_3_option, value); }
        }

        /**
         * whether object uses an automatic line style
         * @return  the autoline field value.
         */
        public bool IsAutoline
        {
            get
            {
                return autoline.IsSet(field_3_option);
            }
            set { field_3_option = autoline.SetShortBoolean(field_3_option, value); }
        }


    }
}

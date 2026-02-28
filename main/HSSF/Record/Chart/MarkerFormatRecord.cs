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

using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// specifies the color, size, and shape of the associated data markers that appear on line, radar, 
    /// and scatter chart groups. The associated data markers are specified by the preceding DataFormat record.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class MarkerFormatRecord : StandardRecord
    {
        public const short sid = 0x1009;
        private int field_1_rgbFore;
        private int field_2_rgbBack;
        private short field_3_imk;
        private short field_4_flag;
        private short field_5_icvFore;
        private short field_6_icvBack;
        private int field_7_miSize;

        BitField fAuto = BitFieldFactory.GetInstance(0x1);
        BitField fNotShowInt = BitFieldFactory.GetInstance(0x10);
        BitField fNotShowBrd = BitFieldFactory.GetInstance(0x20);

        protected override int DataSize
        {
            get { return 4 + 4 + 2 + 2 + 2 + 2 + 4; }
        }

        public MarkerFormatRecord()
        {
        }

        public MarkerFormatRecord(RecordInputStream ris)
        {
            field_1_rgbFore = ris.ReadInt();
            field_2_rgbBack = ris.ReadInt();
            field_3_imk = ris.ReadShort();
            field_4_flag = ris.ReadShort();
            field_5_icvFore = ris.ReadShort();
            field_6_icvBack = ris.ReadShort();
            field_7_miSize = ris.ReadInt();
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_rgbFore);
            out1.WriteInt(field_2_rgbBack);
            out1.WriteShort(field_3_imk);
            out1.WriteShort(field_4_flag);
            out1.WriteShort(field_5_icvFore);
            out1.WriteShort(field_6_icvBack);
            out1.WriteInt(field_7_miSize);
        }
        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[MARKERFORMAT]").AppendLine()
                .Append("   .rgbFore          =").Append(HexDump.ToHex(field_1_rgbFore)).Append("(").Append(field_1_rgbFore).AppendLine(")")
                .Append("   .rgbBack          =").Append(HexDump.ToHex(field_2_rgbBack)).Append("(").Append(field_2_rgbBack).AppendLine(")")
                .Append("   .imk              =").Append(HexDump.ToHex(field_3_imk)).Append("(").Append(field_3_imk).AppendLine(")")
                .Append("   .flag             =").Append(HexDump.ToHex(field_4_flag)).Append("(").Append(field_4_flag).AppendLine(")")
                .Append("       .fAuto        =").Append(this.Auto).AppendLine()
                .Append("       .fNotShowInt  =").Append(this.NotShowInt).AppendLine()
                .Append("       .fNotShowBrd  =").Append(this.NotShowBorder).AppendLine()
                .Append("   .icvFore          =").Append(HexDump.ToHex(field_5_icvFore)).Append("(").Append(field_5_icvFore).AppendLine(")")
                .Append("   .icvBack          =").Append(HexDump.ToHex(field_6_icvBack)).Append("(").Append(field_6_icvBack).AppendLine(")")
                .Append("   .miSize           =").Append(HexDump.ToHex(field_7_miSize)).Append("(").Append(field_7_miSize).AppendLine(")")
                .AppendLine("[/MARKERFORMAT]");
            return buffer.ToString();
        }
        public override object Clone()
        {
            MarkerFormatRecord record = new MarkerFormatRecord();
            record.Auto = this.Auto;
            record.DataMarkerType = this.DataMarkerType;
            record.IcvBack = this.IcvBack;
            record.IcvFore = this.IcvFore;
            record.NotShowBorder = this.NotShowBorder;
            record.NotShowInt = this.NotShowInt;
            record.RGBBack = this.RGBBack;
            record.RGBFore = this.RGBFore;
            record.Size = this.Size;
            return record;
        }
        public override short Sid
        {
            get { return sid; }
        }
        /// <summary>
        /// the border color of the data marker.
        /// </summary>
        public int RGBFore
        {
            get { return field_1_rgbFore; }
            set { field_1_rgbFore = value; }
        }
        /// <summary>
        /// the interior color of the data marker.
        /// </summary>
        public int RGBBack
        {
            get { return field_2_rgbBack; }
            set { field_2_rgbBack = value; }
        }
        /// <summary>
        /// the type of data marker.
        /// </summary>
        public short DataMarkerType
        {
            get { return field_3_imk; }
            set { field_3_imk = value; }
        }
        /// <summary>
        /// whether the data marker is automatically generated.
        /// false The data marker is not automatically generated.
        /// true  The data marker type, size, and color are automatically generated and the values are set accordingly in this record.
        /// </summary>
        public bool Auto
        {
            get { return fAuto.IsSet(field_4_flag); }
            set { field_4_flag = fAuto.SetShortBoolean(field_4_flag, value); }
        }
        /// <summary>
        /// whether to show the data marker interior.
        /// false  The data marker interior is shown.
        /// true   The data marker interior is not shown.
        /// </summary>
        public bool NotShowInt
        {
            get { return fNotShowInt.IsSet(field_4_flag); }
            set { field_4_flag = fNotShowInt.SetShortBoolean(field_4_flag, value); }
        }
        /// <summary>
        /// whether to show the data marker border.
        /// false The data marker border is shown.
        /// true  The data marker border is not shown.
        /// </summary>
        public bool NotShowBorder
        {
            get { return fNotShowBrd.IsSet(field_4_flag); }
            set { field_4_flag = fNotShowBrd.SetShortBoolean(field_4_flag, value); }
        }
        /// <summary>
        /// the border color of the data marker.
        /// </summary>
        public short IcvFore
        {
            get { return field_5_icvFore; }
            set { field_5_icvFore = value; }
        }
        /// <summary>
        /// the interior color of the data marker.
        /// </summary>
        public short IcvBack
        {
            get { return field_6_icvBack; }
            set { field_6_icvBack = value; }
        }
        /// <summary>
        /// specifies the size in twips of the data marker.
        /// </summary>
        public int Size
        {
            get { return field_7_miSize; }
            set { field_7_miSize = value; }
        }
    }

    public enum MarkerType
    {
        None = 0,
        Square = 1,
        DiamondShaped = 2,
        Triangular = 3,
        SquareWithX = 4,
        SquareWithAsterisk = 5,
        ShortBar = 6,
        LongBar = 7,
        Circular = 8,
        SquareWithPlus = 9
    }
}

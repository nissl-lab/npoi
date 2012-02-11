using System;
using System.Collections.Generic;
using System.Text;

using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    public class Chart3DBarShapeRecord:StandardRecord
    {
        public const short sid = 4191;

        byte field_1_riser = 0;
        byte field_2_taper = 0;

        public Chart3DBarShapeRecord()
        { 
        
        }

        public Chart3DBarShapeRecord(RecordInputStream in1)
        {
            field_1_riser = (byte)in1.ReadByte();
            field_2_taper = (byte)in1.ReadByte();
        }

        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[Chart3DBarShape]\n");
            buffer.Append("    .axisType             = ")
                .Append("0x").Append(HexDump.ToHex(Riser))
                .Append(" (").Append(Riser).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .x                    = ")
                .Append("0x").Append(HexDump.ToHex(Taper))
                .Append(" (").Append(Taper).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/Chart3DBarShape]\n");
            return buffer.ToString();
        }
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(field_1_riser);
            out1.WriteByte(field_2_taper);
        }

        protected override int DataSize
        {
            get { return 1 + 1; }
        }

        public override object Clone()
        {
            Chart3DBarShapeRecord record = new Chart3DBarShapeRecord();
            record.Riser = this.Riser;
            record.Taper = this.Taper;
            return record;
        }

        public override short Sid
        {
            get { return sid; }
        }

        public byte Riser
        {
            get { return field_1_riser; }
            set { field_1_riser = value; }
        }

        public byte Taper
        {
            get { return field_2_taper; }
            set { field_2_taper = value; }
        }
    }
}

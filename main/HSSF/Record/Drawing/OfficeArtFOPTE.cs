using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;
using NPOI.SS.UserModel.Drawing;

namespace NPOI.HSSF.Record.Drawing
{
    public class OfficeArtFOPTE
    {
        //private OfficeArtFOPTEOPID field_1_opid;
        //private int field_2_op;
        private byte[] complexData = new byte[0];

        public OfficeArtFOPTE(RecordInputStream ris)
        {
            Opid = new OfficeArtFOPTEOPID((ushort)ris.ReadUShort());
            Op = ris.ReadInt();
            if (Opid.IsComplex)
            {
                complexData = new byte[Op];
                for (int i = 0; i < complexData.Length; i++)
                    complexData[i] = (byte)ris.ReadByte();
            }
        }
        public int DataSize
        {
            get { return 2 + 4 + complexData.Length; }
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Opid.FieldOpid);
            out1.WriteInt(Op);
            out1.Write(complexData);
        }

        public OfficeArtFOPTEOPID Opid
        {
            get;
            set;
        }
        public int Op
        {
            get;
            set;
        }
        public byte[] ComplexData
        {
            get { return complexData; }
        }

        public override string ToString()
        {
            return string.Format("    " + OfficeArtProperties.GetFillStyleName(this.Opid.OpId) + "opid=" + HexDump.ShortToHex(this.Opid.OpId)
                + "; op=" + HexDump.IntToHex(this.Op));
            
        }
    }
}

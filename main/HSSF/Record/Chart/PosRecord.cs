using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    public enum PositionMode : short
    { 
        MDFX = 0x0000,
        MDABS = 0x0001,
        MDPARENT = 0x0002,
        MDKTH = 0x0003,
        MDCHART = 0x0004
    }
    [Obsolete]
    public class PosRecord : StandardRecord
    {

        public PosRecord()
        { 
        
        }

        public const short sid = 0x104F;

        //the positioning mode for the upper-left corner
        private short mdTopLt;
        //the positioning mode for the lower-right corner
        private short mdBotRt;
        private short x1;
        private short y1;
        private short x2;
        private short y2;

        public PosRecord(RecordInputStream in1)
        {
            mdTopLt = in1.ReadShort();
            mdBotRt = in1.ReadShort();
            x1 = in1.ReadShort();
            in1.ReadShort();    //unused1
            y1 = in1.ReadShort();
            in1.ReadShort();    //unused2
            x2 = in1.ReadShort();
            in1.ReadShort();    //unused3
            y2 = in1.ReadShort();
            in1.ReadShort();    //unused4
        }

        protected override int DataSize
        {
            get { return 20; }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(mdTopLt);
            out1.WriteShort(mdBotRt);
            out1.WriteShort(x1);
            out1.WriteShort(0);
            out1.WriteShort(y1);
            out1.WriteShort(0);
            out1.WriteShort(x2);
            out1.WriteShort(0);
            out1.WriteShort(y2);
            out1.WriteShort(0);
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override string ToString()
        {
            return base.ToString();
        }

        public PositionMode MDTopLt
        {
            get 
            {
                return (PositionMode)mdTopLt;
            }
            set 
            {
                mdTopLt = (short)value;
            }
        }

        public PositionMode MdBotRt
        {
            get
            {
                return (PositionMode)mdBotRt;
            }
            set {
                mdBotRt = (short)value;
            }
        }
        public short X1
        {
            get { return x1; }
            set { x1 = value; }
        }
        public short X2
        {
            get { return x2; }
            set { x2 = value; }            
        }
        public short Y1
        {
            get { return y1; }
            set { y1 = value; }
        }
        public short Y2
        {
            get { return y2; }
            set { y2 = value; }
        }
    }
}

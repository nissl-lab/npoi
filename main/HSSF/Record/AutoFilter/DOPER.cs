using System;
using NPOI.Util;

namespace NPOI.HSSF.Record.AutoFilter
{
    public enum DOPERComparisonCode:byte
    { 
        Unknown=0,
        Less=1,
        Equal=2,
        LessThan=3,
        More=4,
        NotEqual=5,
        GreaterThan=6
    }
    public enum DOPERType:byte
    { 
        FilterCondition=0,
        RKNumber=0x02,
        IEEENumber=0x04,
        String = 0x06,
        BooleanOrErrors=0x08,
        MatchAllBlanks=0x0C,
        MatchNoneBlank=0x0E
    }

    public enum DOPERErrorValue : byte
    { 
        NULL=0,     //#NULL!
        DIV0=0x07,  //#DIV/0!
        VALUE=0x0F, //#VALUE!
        REF=0x17,   //#REF!
        NAME=0x1D,  //#NAME?
        NUM=0x24,   //#NUM!
        NA=0x2A     //N/A
    }

    /// <summary>
    /// DOPER Structure for AutoFilter record
    /// </summary>
    /// <remarks>author: Tony Qu</remarks>
    public class DOPERRecord : RecordBase
    {
        private DOPERType vt;
        private byte grbitSgn;
        private RKRecord _RK;
        private double _IEEENumber;
        private byte CCH;
        private byte fError;
        private byte bBoolErr;

        public DOPERRecord()
        { 
        
        }

        public DOPERRecord(RecordInputStream in1)
        {
            vt=(DOPERType)in1.ReadByte();
            switch (vt)
            { 
                case DOPERType.RKNumber:
                    grbitSgn = (byte)in1.ReadByte();
                    _RK = new RKRecord(in1);
                    in1.ReadInt();  //reserved
                    break;
                case DOPERType.IEEENumber:
                    grbitSgn = (byte)in1.ReadByte();
                    _IEEENumber = in1.ReadDouble();
                    break;          
                case DOPERType.String:
                    grbitSgn = (byte)in1.ReadByte();
                    in1.ReadInt();  //reserved
                    CCH = (byte)in1.ReadByte();
                    in1.ReadByte();     //reserved
                    in1.ReadShort();    //reserved
                    break;
                case DOPERType.BooleanOrErrors:
                    grbitSgn = (byte)in1.ReadByte();
                    fError=(byte)in1.ReadByte();
                    bBoolErr=(byte)in1.ReadByte();
                    in1.ReadShort();    //reserved
                    in1.ReadInt();      //reserved
                    break;
                default:    //FilterCondition,MatchAllBlanks,MatchNoneBlank
                    grbitSgn = 0;
                    in1.ReadByte();    //reserved
                    in1.ReadLong();    //reserved
                    break;
            }
        }
        public override int RecordSize
        {
            get
            {
                return 10;
            }
        }
        public int Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte((byte)vt);
            switch (vt)
            {
                case DOPERType.RKNumber:
                    out1.WriteByte(grbitSgn);
                    _RK.Serialize(out1);
                    out1.WriteInt(0);  //reserved
                    break;
                case DOPERType.IEEENumber:
                    out1.WriteByte(grbitSgn);
                    out1.WriteDouble(_IEEENumber);
                    break;
                case DOPERType.String:
                    out1.WriteByte(grbitSgn);
                    out1.WriteInt(0);  //reserved
                    out1.WriteByte(CCH);
                    out1.WriteByte(0);      //reserved
                    out1.WriteShort(0);     //reserved
                    break;
                case DOPERType.BooleanOrErrors:
                    out1.WriteByte(grbitSgn);
                    out1.WriteByte(fError);
                    out1.WriteByte(bBoolErr);
                    out1.WriteShort(0);    //reserved
                    out1.WriteInt(0);      //reserved
                    break;
                default:    //FilterCondition,MatchAllBlanks,MatchNoneBlank
                    out1.WriteByte(0);      //reserved
                    out1.WriteLong(0);      //reserved
                    break;
            }
            return this.RecordSize;            
        }
        public virtual Record CloneViaReserialise()
        {
            throw new NotImplementedException("Please implement it in subclass");
        }
        public override int Serialize(int offset, byte[] data)
        {
            LittleEndianByteArrayOutputStream out1 = new LittleEndianByteArrayOutputStream(data, offset,this.RecordSize);
            int result=this.Serialize(out1);
            if (out1.WriteIndex - offset != this.RecordSize)
            {
                throw new InvalidOperationException("Error in serialization of (" + this.GetType().Name + "): "
                        + "Incorrect number of bytes written - expected "
                        + this.RecordSize + " but got " + (out1.WriteIndex - offset));
            }
            return result;
        }

        public DOPERType DataType
        {
            get { return vt; }
            set { vt = value; }
        }

        public DOPERComparisonCode ComparisonCode
        {
            get { return (DOPERComparisonCode)grbitSgn;}
            set { grbitSgn = (byte)value; }
        }

        public Double IEEENumber
        {
            get { return _IEEENumber; }
            set { 
                _IEEENumber = value;
                vt = DOPERType.IEEENumber;
            }
        }
        /// <summary>
        /// get or set the RK record
        /// </summary>
        public RKRecord RK
        {
            get { return _RK; }
            set { 
                _RK = value;
                vt = DOPERType.RKNumber;
            }
        }
        /// <summary>
        /// Gets or sets Length of the string (the string is stored in the rgch field that follows the DOPER structures)
        /// </summary>
        public byte LengthOfString
        {
            get { return CCH; }
            set {
                if (value > 252)
                    throw new ArgumentOutOfRangeException("The length of string must be less than or equal to 252");
                CCH = value;
                vt = DOPERType.String;
            }
        }
        /// <summary>
        /// Whether the bBoolErr field contains a Boolean value
        /// </summary>
        public bool IsBooleanValue
        {
            get { return fError == 0; }
        }
        /// <summary>
        /// Whether the bBoolErr field contains a Error value
        /// </summary>
        public bool IsErrorValue
        {
            get { return fError == 1; }
        }
        /// <summary>
        /// Get or sets the boolean value
        /// </summary>
        public bool BooleanValue
        {
            get { return bBoolErr == 1; }
            set {
                bBoolErr = value ? (byte)1 : (byte)0;
                fError = 0;
                vt = DOPERType.BooleanOrErrors;
            }
        }
        /// <summary>
        /// Get or sets the boolean value
        /// </summary>
        public DOPERErrorValue ErrorValue
        {
            get {
                 return (DOPERErrorValue)bBoolErr; 
            }
            set {
                bBoolErr = (byte)value;
                fError = 1;
                vt = DOPERType.BooleanOrErrors;
            }
        }
    }
}

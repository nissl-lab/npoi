using System;
using NPOI.Util;
using System.Globalization;

namespace NPOI.HSSF.Record
{
    /// <summary>
    /// FtSbs structure
    /// </summary>
    public class ScrollableObjectSubRecord : SubRecord
    {
        private short field_1_iVal = 0;
        private short field_2_iMin = 0;
        private short field_3_iMax = 0;
        private short field_4_dInc = 0;
        private short field_5_dPage = 0;
        private short field_6_fHoriz = 0;
        private short field_7_dxScroll = 0;
        private short field_8_options = 0;
        private BitField fDrawFlag = BitFieldFactory.GetInstance(0x01);
        private BitField fDrawSliderOnly = BitFieldFactory.GetInstance(0x02);
        private BitField fTrackElevator = BitFieldFactory.GetInstance(0x04);
        private BitField fNo3d = BitFieldFactory.GetInstance(0x08);

        public ScrollableObjectSubRecord()
        { 
        
        }

        public ScrollableObjectSubRecord(ILittleEndianInput in1, int size)
        {
            if (size !=this.DataSize)
            {
                throw new RecordFormatException(string.Format(CultureInfo.CurrentCulture, "Expected size {0} but got ({1})", this.DataSize, size));
            }
            in1.ReadInt();
            field_1_iVal=in1.ReadShort();
            field_2_iMin=in1.ReadShort();
            field_3_iMax=in1.ReadShort();
            field_4_dInc=in1.ReadShort();
            field_5_dPage=in1.ReadShort();
            field_6_fHoriz = in1.ReadShort();
            field_7_dxScroll = in1.ReadShort();
            field_8_options = in1.ReadShort();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(DataSize);

            out1.WriteInt(0);
            out1.WriteShort(field_1_iVal);
            out1.WriteShort(field_2_iMin);
            out1.WriteShort(field_3_iMax);
            out1.WriteShort(field_4_dInc);
            out1.WriteShort(field_5_dPage);
            out1.WriteShort(field_6_fHoriz);
            out1.WriteShort(field_7_dxScroll);
            out1.WriteShort(field_8_options);
        }

        public override int DataSize
        {
            get { return 20; }
        }
        public const short sid = 0x0C;
        public override short Sid
        {
            get { return sid; }
        }

        public short CurrentValue
        {
            get { return field_1_iVal; }
            set {
                if (field_1_iVal < field_2_iMin || field_1_iVal > field_3_iMax)
                    throw new ArgumentOutOfRangeException("invalid value");
                field_1_iVal = value; 
            }
        }
        public short MaxValue
        {
            get { return field_2_iMin; }
            set { field_2_iMin = value; }
        }
        public short MinValue
        {
            get { return field_3_iMax; }
            set { field_3_iMax = value; }
        }
        public short IncreaseAmountChanged
        {
            get { return field_4_dInc; }
            set { field_4_dInc = value; }
        }
        public short PageAmountChanged
        {
            get { return field_5_dPage; }
            set { field_5_dPage = value; }
        }
        public bool IsHorizontal
        {
            get { return field_6_fHoriz==1; }
            set { field_6_fHoriz = value ? (short)1 : (short)0; }
        }
        public short ScrollbarWidthInPixel
        {
            get { return field_7_dxScroll; }
            set { field_7_dxScroll = value; }
        }
        public bool IsVisible
        {
            get { return fDrawFlag.IsSet(field_8_options); }
            set { field_8_options=fDrawFlag.SetShortBoolean(field_8_options, value);  }
        }
        public bool IsOnlySilderPortionVisible
        {
            get { return fDrawSliderOnly.IsSet(field_8_options); }
            set { field_8_options = fDrawSliderOnly.SetShortBoolean(field_8_options, value); }
        }
        public bool IsTrackElevator
        {
            get { return fTrackElevator.IsSet(field_8_options); }
            set { field_8_options = fTrackElevator.SetShortBoolean(field_8_options, value); }            
        }
        public bool IsNo3D
        {
            get { return fNo3d.IsSet(field_8_options); }
            set { field_8_options = fNo3d.SetShortBoolean(field_8_options, value); }                    
        }
        public override object Clone()
        {
            ScrollableObjectSubRecord rec = new ScrollableObjectSubRecord();
            rec.field_1_iVal = field_1_iVal;
            rec.field_2_iMin = field_2_iMin;
            rec.field_3_iMax = field_3_iMax;
            rec.field_4_dInc = field_4_dInc;
            rec.field_5_dPage = field_5_dPage;
            rec.field_6_fHoriz = field_6_fHoriz;
            rec.field_7_dxScroll = field_7_dxScroll;
            rec.field_8_options = field_8_options;
            return rec;
        }
    }
}

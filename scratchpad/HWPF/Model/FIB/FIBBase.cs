using System;
using System.Text;
using NPOI.Util;

namespace NPOI.HWPF.Model
{
    public class FIBBase
    {
        short field_1_wIdent;
        short field_2_nFib;
        short field_4_lid;
        short field_5_pnNext;
        short field_6_flags=0;

        static BitField field_6_fDot = BitFieldFactory.GetInstance(0x01);
        static BitField field_6_fGlsy = BitFieldFactory.GetInstance(0x02);
        static BitField field_6_fComplex = BitFieldFactory.GetInstance(0x04);
        static BitField field_6_fHasPic = BitFieldFactory.GetInstance(0x08);
        static BitField field_6_cQuickSaves = BitFieldFactory.GetInstance(0xF0);
        static BitField field_6_fEncrypted = BitFieldFactory.GetInstance(0x100);
        static BitField field_6_fWhichTblStm = BitFieldFactory.GetInstance(0x200);
        static BitField field_6_fReadOnlyRecommended = BitFieldFactory.GetInstance(0x400);
        static BitField field_6_fWriteReservation = BitFieldFactory.GetInstance(0x800);
        static BitField field_6_fExtChar = BitFieldFactory.GetInstance(0x1000);
        static BitField field_6_fLoadOverride = BitFieldFactory.GetInstance(0x2000);
        static BitField field_6_fFarEast = BitFieldFactory.GetInstance(0x4000);
        static BitField field_6_fObfuscated = BitFieldFactory.GetInstance(0x8000);

        short field_7_nFibBack;
        int field_8_lKey;
        //byte field_9_envr;
        byte field_10_flags=0;

        //BitField field_10_fMac = BitFieldFactory.GetInstance(0x01);
        //BitField field_10_fEmptySpecial = BitFieldFactory.GetInstance(0x02);
        BitField field_10_fLoadOverridePage = BitFieldFactory.GetInstance(0x04);



        public FIBBase()
        {

        }

        public void Deserialize(HWPFStream stream)
        {
            field_1_wIdent = stream.ReadShort();
            field_2_nFib = stream.ReadShort();
            stream.ReadShort();
            field_4_lid = stream.ReadShort();
            field_5_pnNext = stream.ReadShort();
            field_6_flags = stream.ReadShort();
            field_7_nFibBack = stream.ReadShort();
            field_8_lKey = stream.ReadInt();
            stream.ReadByte();  //field 9
            field_10_flags = (byte)stream.ReadByte();
            stream.ReadShort(); //reserved3
            stream.ReadShort(); //reserved4
            stream.ReadInt();   //reserved5
            stream.ReadInt();   //reserved6
        }

        public short wIdent
        {
            get { return field_1_wIdent; }
            set { field_1_wIdent = value; }
        }

        public short nFib
        {
            get { return field_2_nFib; }
            set { field_2_nFib = value; }
        }

        public short lid
        {
            get { return field_4_lid; }
            set { field_4_lid = value; }
        }

        public short pnNext
        {
            get { return field_5_pnNext; }
            set { field_5_pnNext = value; }
        }

        public bool IsDocumentTemplate
        {
            get { return field_6_fDot.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fDot.SetShortBoolean(field_6_flags, value); }
        }

        public bool IsComplex
        {
            get { return field_6_fComplex.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fComplex.SetShortBoolean(field_6_flags, value); }            
        }

        public bool HasPicture
        {
            get { return field_6_fHasPic.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fHasPic.SetShortBoolean(field_6_flags, value); }                        
        }

        public short QuickSavesTimes
        {
            get { return field_6_cQuickSaves.GetShortValue(field_6_flags); }
            set { field_6_flags = field_6_cQuickSaves.SetShortValue(field_6_flags, value); }
        }

        public bool IsEncrypted
        {
            get { return field_6_fEncrypted.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fEncrypted.SetShortBoolean(field_6_flags, value); }            
        }

        public short WhichTableStream
        {
            get { return field_6_fWhichTblStm.GetShortValue(field_6_flags); }
            set { field_6_flags = field_6_fWhichTblStm.SetShortValue(field_6_flags, value); }
        }

        public bool IsReadOnlyRecommended
        {
            get { return field_6_fReadOnlyRecommended.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fReadOnlyRecommended.SetShortBoolean(field_6_flags, value); }                  
        }

        public bool HasWriteReservationPassword
        {
            get { return field_6_fWriteReservation.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fWriteReservation.SetShortBoolean(field_6_flags, value); }           
        }

        public bool IsLoadOverride
        {
            get { return field_6_fLoadOverride.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fLoadOverride.SetShortBoolean(field_6_flags, value); }
        }

        public bool IsFarEast
        {
            get { return field_6_fFarEast.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fFarEast.SetShortBoolean(field_6_flags, value); }            
        }

        public bool UseXORObfuscation
        {
            get { return field_6_fObfuscated.IsSet(field_6_flags); }
            set { field_6_flags = field_6_fObfuscated.SetShortBoolean(field_6_flags, value); }              
        }

        public short nFibBack
        {
            get { return field_7_nFibBack; }
            set { field_7_nFibBack = value; }
        }

        public int lKey
        {
            get { return field_8_lKey; }
            set { field_8_lKey = value; }
        }

        public bool IsLoadOverridePage
        {
            get { return field_10_fLoadOverridePage.IsSet(field_10_flags); }
            set { field_10_flags = field_10_fLoadOverridePage.SetByteBoolean(field_10_flags, value); }              
        }

        public int fExtChar
        {
            get { return field_6_fExtChar.GetValue(field_6_flags); }
        }
    }
}

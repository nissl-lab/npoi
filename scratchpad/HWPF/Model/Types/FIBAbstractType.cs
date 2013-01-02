
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */



namespace NPOI.HWPF.Model.Types
{

    using System;
    using NPOI.Util;

    using NPOI.HWPF.UserModel;
    using System.Text;
    /**
     * Base part of the File information Block (FibBase). Holds the core part of the FIB, from the first 32 bytes.
     * NOTE: This source Is automatically generated please do not modify this file.  Either subclass or
     *       remove the record in src/records/definitions.

     * @author Andrew C. Oliver
     */
    public abstract class FIBAbstractType : BaseObject
    {

        protected int field_1_wIdent;
        protected int field_2_nFib;
        protected int field_3_nProduct;
        protected int field_4_lid;
        protected int field_5_pnNext;
        protected short field_6_options;
        private static BitField fDot = BitFieldFactory.GetInstance(0x0001);
        private static BitField fGlsy = BitFieldFactory.GetInstance(0x0002);
        private static BitField fComplex = BitFieldFactory.GetInstance(0x0004);
        private static BitField fHasPic = BitFieldFactory.GetInstance(0x0008);
        private static BitField cQuickSaves = BitFieldFactory.GetInstance(0x00F0);
        private static BitField fEncrypted = BitFieldFactory.GetInstance(0x0100);
        private static BitField fWhichTblStm = BitFieldFactory.GetInstance(0x0200);
        private static BitField fReadOnlyRecommended = BitFieldFactory.GetInstance(0x0400);
        private static BitField fWriteReservation = BitFieldFactory.GetInstance(0x0800);
        private static BitField fExtChar = BitFieldFactory.GetInstance(0x1000);
        private static BitField fLoadOverride = BitFieldFactory.GetInstance(0x2000);
        private static BitField fFarEast = BitFieldFactory.GetInstance(0x4000);
        private static BitField fCrypto = BitFieldFactory.GetInstance(0x8000);
        protected int field_7_nFibBack;
        protected int field_8_lKey;
        protected int field_9_envr;
        protected short field_10_history;
        private static BitField fMac = BitFieldFactory.GetInstance(0x0001);
        private static BitField fEmptySpecial = BitFieldFactory.GetInstance(0x0002);
        private static BitField fLoadOverridePage = BitFieldFactory.GetInstance(0x0004);
        private static BitField fFutureSavedUndo = BitFieldFactory.GetInstance(0x0008);
        private static BitField fWord97Saved = BitFieldFactory.GetInstance(0x0010);
        private static BitField fSpare0 = BitFieldFactory.GetInstance(0x00FE);
        protected int field_11_chs;       /** Latest docs say this Is Reserved3! */
        protected int field_12_chsTables; /** Latest docs say this Is Reserved4! */
        protected int field_13_fcMin;     /** Latest docs say this Is Reserved5! */
        protected int field_14_fcMac;     /** Latest docs say this Is Reserved6! */


        public FIBAbstractType()
        {

        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_wIdent = LittleEndian.GetShort(data, 0x0 + offset);
            field_2_nFib = LittleEndian.GetShort(data, 0x2 + offset);
            field_3_nProduct = LittleEndian.GetShort(data, 0x4 + offset);
            field_4_lid = LittleEndian.GetShort(data, 0x6 + offset);
            field_5_pnNext = LittleEndian.GetShort(data, 0x8 + offset);
            field_6_options = LittleEndian.GetShort(data, 0xa + offset);
            field_7_nFibBack = LittleEndian.GetShort(data, 0xc + offset);
            field_8_lKey = LittleEndian.GetShort(data, 0xe + offset);
            field_9_envr = LittleEndian.GetShort(data, 0x10 + offset);
            field_10_history = LittleEndian.GetShort(data, 0x12 + offset);
            field_11_chs = LittleEndian.GetShort(data, 0x14 + offset);
            field_12_chsTables = LittleEndian.GetShort(data, 0x16 + offset);
            field_13_fcMin = LittleEndian.GetInt(data, 0x18 + offset);
            field_14_fcMac = LittleEndian.GetInt(data, 0x1c + offset);

        }

        public void Serialize(byte[] data, int offset)
        {
            LittleEndian.PutShort(data, 0x0 + offset, (short)field_1_wIdent); ;
            LittleEndian.PutShort(data, 0x2 + offset, (short)field_2_nFib); ;
            LittleEndian.PutShort(data, 0x4 + offset, (short)field_3_nProduct); ;
            LittleEndian.PutShort(data, 0x6 + offset, (short)field_4_lid); ;
            LittleEndian.PutShort(data, 0x8 + offset, (short)field_5_pnNext); ;
            LittleEndian.PutShort(data, 0xa + offset, (short)field_6_options); ;
            LittleEndian.PutShort(data, 0xc + offset, (short)field_7_nFibBack); ;
            LittleEndian.PutShort(data, 0xe + offset, (short)field_8_lKey); ;
            LittleEndian.PutShort(data, 0x10 + offset, (short)field_9_envr); ;
            LittleEndian.PutShort(data, 0x12 + offset, (short)field_10_history); ;
            LittleEndian.PutShort(data, 0x14 + offset, (short)field_11_chs); ;
            LittleEndian.PutShort(data, 0x16 + offset, (short)field_12_chsTables); ;
            LittleEndian.PutInt(data, 0x18 + offset, field_13_fcMin); ;
            LittleEndian.PutInt(data, 0x1c + offset, field_14_fcMac); ;

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FIB]\n");

            buffer.Append("    .wIdent               = ");
            buffer.Append(" (").Append(GetWIdent()).Append(" )\n");

            buffer.Append("    .nFib                 = ");
            buffer.Append(" (").Append(GetNFib()).Append(" )\n");

            buffer.Append("    .nProduct             = ");
            buffer.Append(" (").Append(GetNProduct()).Append(" )\n");

            buffer.Append("    .lid                  = ");
            buffer.Append(" (").Append(GetLid()).Append(" )\n");

            buffer.Append("    .pnNext               = ");
            buffer.Append(" (").Append(GetPnNext()).Append(" )\n");

            buffer.Append("    .options              = ");
            buffer.Append(" (").Append(GetOptions()).Append(" )\n");
            buffer.Append("         .fDot                     = ").Append(IsFDot()).Append('\n');
            buffer.Append("         .fGlsy                    = ").Append(IsFGlsy()).Append('\n');
            buffer.Append("         .fComplex                 = ").Append(IsFComplex()).Append('\n');
            buffer.Append("         .fHasPic                  = ").Append(IsFHasPic()).Append('\n');
            buffer.Append("         .cQuickSaves              = ").Append(GetCQuickSaves()).Append('\n');
            buffer.Append("         .fEncrypted               = ").Append(IsFEncrypted()).Append('\n');
            buffer.Append("         .fWhichTblStm             = ").Append(IsFWhichTblStm()).Append('\n');
            buffer.Append("         .fReadOnlyRecommended     = ").Append(IsFReadOnlyRecommended()).Append('\n');
            buffer.Append("         .fWriteReservation        = ").Append(IsFWriteReservation()).Append('\n');
            buffer.Append("         .fExtChar                 = ").Append(IsFExtChar()).Append('\n');
            buffer.Append("         .fLoadOverride            = ").Append(IsFLoadOverride()).Append('\n');
            buffer.Append("         .fFarEast                 = ").Append(IsFFarEast()).Append('\n');
            buffer.Append("         .fCrypto                  = ").Append(IsFCrypto()).Append('\n');

            buffer.Append("    .nFibBack             = ");
            buffer.Append(" (").Append(GetNFibBack()).Append(" )\n");

            buffer.Append("    .lKey                 = ");
            buffer.Append(" (").Append(GetLKey()).Append(" )\n");

            buffer.Append("    .envr                 = ");
            buffer.Append(" (").Append(GetEnvr()).Append(" )\n");

            buffer.Append("    .history              = ");
            buffer.Append(" (").Append(GetHistory()).Append(" )\n");
            buffer.Append("         .fMac                     = ").Append(IsFMac()).Append('\n');
            buffer.Append("         .fEmptySpecial            = ").Append(IsFEmptySpecial()).Append('\n');
            buffer.Append("         .fLoadOverridePage        = ").Append(IsFLoadOverridePage()).Append('\n');
            buffer.Append("         .fFutureSavedUndo         = ").Append(IsFFutureSavedUndo()).Append('\n');
            buffer.Append("         .fWord97Saved             = ").Append(IsFWord97Saved()).Append('\n');
            buffer.Append("         .fSpare0                  = ").Append(GetFSpare0()).Append('\n');

            buffer.Append("    .chs                  = ");
            buffer.Append(" (").Append(GetChs()).Append(" )\n");

            buffer.Append("    .chsTables            = ");
            buffer.Append(" (").Append(GetChsTables()).Append(" )\n");

            buffer.Append("    .fcMin                = ");
            buffer.Append(" (").Append(GetFcMin()).Append(" )\n");

            buffer.Append("    .fcMac                = ");
            buffer.Append(" (").Append(GetFcMac()).Append(" )\n");

            buffer.Append("[/FIB]\n");
            return buffer.ToString();
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public virtual int GetSize()
        {
            return 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2 + 4 + 4;
        }



        /**
         * Get the wIdent field for the FIB record.
         */
        public int GetWIdent()
        {
            return field_1_wIdent;
        }

        /**
         * Set the wIdent field for the FIB record.
         */
        public void SetWIdent(int field_1_wIdent)
        {
            this.field_1_wIdent = field_1_wIdent;
        }

        /**
         * Get the nFib field for the FIB record.
         */
        public int GetNFib()
        {
            return field_2_nFib;
        }

        /**
         * Set the nFib field for the FIB record.
         */
        public void SetNFib(int field_2_nFib)
        {
            this.field_2_nFib = field_2_nFib;
        }

        /**
         * Get the nProduct field for the FIB record.
         */
        public int GetNProduct()
        {
            return field_3_nProduct;
        }

        /**
         * Set the nProduct field for the FIB record.
         */
        public void SetNProduct(int field_3_nProduct)
        {
            this.field_3_nProduct = field_3_nProduct;
        }

        /**
         * Get the lid field for the FIB record.
         */
        public int GetLid()
        {
            return field_4_lid;
        }

        /**
         * Set the lid field for the FIB record.
         */
        public void SetLid(int field_4_lid)
        {
            this.field_4_lid = field_4_lid;
        }

        /**
         * Get the pnNext field for the FIB record.
         */
        public int GetPnNext()
        {
            return field_5_pnNext;
        }

        /**
         * Set the pnNext field for the FIB record.
         */
        public void SetPnNext(int field_5_pnNext)
        {
            this.field_5_pnNext = field_5_pnNext;
        }

        /**
         * Get the options field for the FIB record.
         */
        public short GetOptions()
        {
            return field_6_options;
        }

        /**
         * Set the options field for the FIB record.
         */
        public void SetOptions(short field_6_options)
        {
            this.field_6_options = field_6_options;
        }

        /**
         * Get the nFibBack field for the FIB record.
         */
        public int GetNFibBack()
        {
            return field_7_nFibBack;
        }

        /**
         * Set the nFibBack field for the FIB record.
         */
        public void SetNFibBack(int field_7_nFibBack)
        {
            this.field_7_nFibBack = field_7_nFibBack;
        }

        /**
         * Get the lKey field for the FIB record.
         */
        public int GetLKey()
        {
            return field_8_lKey;
        }

        /**
         * Set the lKey field for the FIB record.
         */
        public void SetLKey(int field_8_lKey)
        {
            this.field_8_lKey = field_8_lKey;
        }

        /**
         * Get the envr field for the FIB record.
         */
        public int GetEnvr()
        {
            return field_9_envr;
        }

        /**
         * Set the envr field for the FIB record.
         */
        public void SetEnvr(int field_9_envr)
        {
            this.field_9_envr = field_9_envr;
        }

        /**
         * Get the history field for the FIB record.
         */
        public short GetHistory()
        {
            return field_10_history;
        }

        /**
         * Set the history field for the FIB record.
         */
        public void SetHistory(short field_10_history)
        {
            this.field_10_history = field_10_history;
        }

        /**
         * Get the chs field for the FIB record.
         */
        public int GetChs()
        {
            return field_11_chs;
        }

        /**
         * Set the chs field for the FIB record.
         */
        public void SetChs(int field_11_chs)
        {
            this.field_11_chs = field_11_chs;
        }

        /**
         * Get the chsTables field for the FIB record.
         */
        public int GetChsTables()
        {
            return field_12_chsTables;
        }

        /**
         * Set the chsTables field for the FIB record.
         */
        public void SetChsTables(int field_12_chsTables)
        {
            this.field_12_chsTables = field_12_chsTables;
        }

        /**
         * Get the fcMin field for the FIB record.
         */
        public int GetFcMin()
        {
            return field_13_fcMin;
        }

        /**
         * Set the fcMin field for the FIB record.
         */
        public void SetFcMin(int field_13_fcMin)
        {
            this.field_13_fcMin = field_13_fcMin;
        }

        /**
         * Get the fcMac field for the FIB record.
         */
        public int GetFcMac()
        {
            return field_14_fcMac;
        }

        /**
         * Set the fcMac field for the FIB record.
         */
        public void SetFcMac(int field_14_fcMac)
        {
            this.field_14_fcMac = field_14_fcMac;
        }

        /**
         * Sets the fDot field value.
         *
         */
        public void SetFDot(bool value)
        {
            field_6_options = (short)fDot.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fDot field value.
         */
        public bool IsFDot()
        {
            return fDot.IsSet(field_6_options);

        }

        /**
         * Sets the fGlsy field value.
         *
         */
        public void SetFGlsy(bool value)
        {
            field_6_options = (short)fGlsy.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fGlsy field value.
         */
        public bool IsFGlsy()
        {
            return fGlsy.IsSet(field_6_options);

        }

        /**
         * Sets the fComplex field value.
         *
         */
        public void SetFComplex(bool value)
        {
            field_6_options = (short)fComplex.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fComplex field value.
         */
        public bool IsFComplex()
        {
            return fComplex.IsSet(field_6_options);

        }

        /**
         * Sets the fHasPic field value.
         *
         */
        public void SetFHasPic(bool value)
        {
            field_6_options = (short)fHasPic.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fHasPic field value.
         */
        public bool IsFHasPic()
        {
            return fHasPic.IsSet(field_6_options);

        }

        /**
         * Sets the cQuickSaves field value.
         *
         */
        public void SetCQuickSaves(byte value)
        {
            field_6_options = (short)cQuickSaves.SetValue(field_6_options, value);


        }

        /**
         *
         * @return  the cQuickSaves field value.
         */
        public byte GetCQuickSaves()
        {
            return (byte)cQuickSaves.GetValue(field_6_options);

        }

        /**
         * Sets the fEncrypted field value.
         *
         */
        public void SetFEncrypted(bool value)
        {
            field_6_options = (short)fEncrypted.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fEncrypted field value.
         */
        public bool IsFEncrypted()
        {
            return fEncrypted.IsSet(field_6_options);

        }

        /**
         * Sets the fWhichTblStm field value.
         *
         */
        public void SetFWhichTblStm(bool value)
        {
            field_6_options = (short)fWhichTblStm.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fWhichTblStm field value.
         */
        public bool IsFWhichTblStm()
        {
            return fWhichTblStm.IsSet(field_6_options);

        }

        /**
         * Sets the fReadOnlyRecommended field value.
         *
         */
        public void SetFReadOnlyRecommended(bool value)
        {
            field_6_options = (short)fReadOnlyRecommended.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fReadOnlyRecommended field value.
         */
        public bool IsFReadOnlyRecommended()
        {
            return fReadOnlyRecommended.IsSet(field_6_options);

        }

        /**
         * Sets the fWriteReservation field value.
         *
         */
        public void SetFWriteReservation(bool value)
        {
            field_6_options = (short)fWriteReservation.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fWriteReservation field value.
         */
        public bool IsFWriteReservation()
        {
            return fWriteReservation.IsSet(field_6_options);

        }

        /**
         * Sets the fExtChar field value.
         *
         */
        public void SetFExtChar(bool value)
        {
            field_6_options = (short)fExtChar.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fExtChar field value.
         */
        public bool IsFExtChar()
        {
            return fExtChar.IsSet(field_6_options);

        }

        /**
         * Sets the fLoadOverride field value.
         *
         */
        public void SetFLoadOverride(bool value)
        {
            field_6_options = (short)fLoadOverride.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fLoadOverride field value.
         */
        public bool IsFLoadOverride()
        {
            return fLoadOverride.IsSet(field_6_options);

        }

        /**
         * Sets the fFarEast field value.
         *
         */
        public void SetFFarEast(bool value)
        {
            field_6_options = (short)fFarEast.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fFarEast field value.
         */
        public bool IsFFarEast()
        {
            return fFarEast.IsSet(field_6_options);

        }

        /**
         * Sets the fCrypto field value.
         *
         */
        public void SetFCrypto(bool value)
        {
            field_6_options = (short)fCrypto.SetBoolean(field_6_options, value);


        }

        /**
         *
         * @return  the fCrypto field value.
         */
        public bool IsFCrypto()
        {
            return fCrypto.IsSet(field_6_options);

        }

        /**
         * Sets the fMac field value.
         *
         */
        public void SetFMac(bool value)
        {
            field_10_history = (short)fMac.SetBoolean(field_10_history, value);


        }

        /**
         *
         * @return  the fMac field value.
         */
        public bool IsFMac()
        {
            return fMac.IsSet(field_10_history);

        }

        /**
         * Sets the fEmptySpecial field value.
         *
         */
        public void SetFEmptySpecial(bool value)
        {
            field_10_history = (short)fEmptySpecial.SetBoolean(field_10_history, value);


        }

        /**
         *
         * @return  the fEmptySpecial field value.
         */
        public bool IsFEmptySpecial()
        {
            return fEmptySpecial.IsSet(field_10_history);

        }

        /**
         * Sets the fLoadOverridePage field value.
         *
         */
        public void SetFLoadOverridePage(bool value)
        {
            field_10_history = (short)fLoadOverridePage.SetBoolean(field_10_history, value);


        }

        /**
         *
         * @return  the fLoadOverridePage field value.
         */
        public bool IsFLoadOverridePage()
        {
            return fLoadOverridePage.IsSet(field_10_history);

        }

        /**
         * Sets the fFutureSavedUndo field value.
         *
         */
        public void SetFFutureSavedUndo(bool value)
        {
            field_10_history = (short)fFutureSavedUndo.SetBoolean(field_10_history, value);


        }

        /**
         *
         * @return  the fFutureSavedUndo field value.
         */
        public bool IsFFutureSavedUndo()
        {
            return fFutureSavedUndo.IsSet(field_10_history);

        }

        /**
         * Sets the fWord97Saved field value.
         *
         */
        public void SetFWord97Saved(bool value)
        {
            field_10_history = (short)fWord97Saved.SetBoolean(field_10_history, value);


        }

        /**
         *
         * @return  the fWord97Saved field value.
         */
        public bool IsFWord97Saved()
        {
            return fWord97Saved.IsSet(field_10_history);

        }

        /**
         * Sets the fSpare0 field value.
         *
         */
        public void SetFSpare0(byte value)
        {
            field_10_history = (short)fSpare0.SetValue(field_10_history, value);


        }

        /**
         *
         * @return  the fSpare0 field value.
         */
        public byte GetFSpare0()
        {
            return (byte)fSpare0.GetValue(field_10_history);

        }


    }
}
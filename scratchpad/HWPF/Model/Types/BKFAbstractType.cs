/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.Model.Types
{

    using NPOI.Util;
    using System.Text;
    using System;

    /**
     * BooKmark First descriptor (BKF).
     * <p>
     * Class and fields descriptions are quoted from Microsoft Office Word 97-2007
     * Binary File Format (.doc) Specification
     * 
     * NOTE: This source is automatically generated please do not modify this file.
     * Either subclass or remove the record in src/types/defInitions.
     * 
     * @author Sergey Vladimirov; according to Microsoft Office Word 97-2007 Binary
     *         File Format (.doc) Specification
     */

    public abstract class BKFAbstractType : BaseObject
    {

        protected short field_1_ibkl;
        protected short field_2_bkf_flags;
        private static BitField itcFirst = new BitField(0x007F);
        private static BitField fPub = new BitField(0x0080);
        private static BitField itcLim = new BitField(0x7F00);
        private static BitField fCol = new BitField(0x8000);

        protected BKFAbstractType()
        {
        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_ibkl = LittleEndian.GetShort(data, 0x0 + offset);
            field_2_bkf_flags = LittleEndian.GetShort(data, 0x2 + offset);
        }

        public void Serialize(byte[] data, int offset)
        {
            LittleEndian.PutShort(data, 0x0 + offset, field_1_ibkl);
            LittleEndian.PutShort(data, 0x2 + offset, field_2_bkf_flags);
        }

        /**
         * Size of record
         */
        public static int GetSize()
        {
            return 0 + 2 + 2;
        }

        public override String ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("[BKF]\n");
            builder.Append("    .ibkl                 = ");
            builder.Append(" (").Append(GetIbkl()).Append(" )\n");
            builder.Append("    .bkf_flags            = ");
            builder.Append(" (").Append(GetBkf_flags()).Append(" )\n");
            builder.Append("         .itcFirst                 = ").Append(GetItcFirst()).Append('\n');
            builder.Append("         .fPub                     = ").Append(IsFPub()).Append('\n');
            builder.Append("         .itcLim                   = ").Append(GetItcLim()).Append('\n');
            builder.Append("         .fCol                     = ").Append(IsFCol()).Append('\n');

            builder.Append("[/BKF]\n");
            return builder.ToString();
        }

        /**
         * Index to BKL entry in plcfbkl that describes the ending position of this bookmark in the CP stream.
         */
        public short GetIbkl()
        {
            return field_1_ibkl;
        }

        /**
         * Index to BKL entry in plcfbkl that describes the ending position of this bookmark in the CP stream.
         */
        public void SetIbkl(short field_1_ibkl)
        {
            this.field_1_ibkl = field_1_ibkl;
        }

        /**
         * Get the bkf_flags field for the BKF record.
         */
        public short GetBkf_flags()
        {
            return field_2_bkf_flags;
        }

        /**
         * Set the bkf_flags field for the BKF record.
         */
        public void SetBkf_flags(short field_2_bkf_flags)
        {
            this.field_2_bkf_flags = field_2_bkf_flags;
        }

        /**
         * Sets the itcFirst field value.
         * When bkf.fCol==1, this is the index to the first column of a table column bookmark
         */
        public void SetItcFirst(byte value)
        {
            field_2_bkf_flags = (short)itcFirst.SetValue(field_2_bkf_flags, value);
        }

        /**
         * When bkf.fCol==1, this is the index to the first column of a table column bookmark
         * @return  the itcFirst field value.
         */
        public byte GetItcFirst()
        {
            return (byte)itcFirst.GetValue(field_2_bkf_flags);
        }

        /**
         * Sets the fPub field value.
         * When 1, this indicates that this bookmark is marking the range of a Macintosh Publisher section
         */
        public void SetFPub(bool value)
        {
            field_2_bkf_flags = (short)fPub.SetBoolean(field_2_bkf_flags, value);
        }

        /**
         * When 1, this indicates that this bookmark is marking the range of a Macintosh Publisher section
         * @return  the fPub field value.
         */
        public bool IsFPub()
        {
            return fPub.IsSet(field_2_bkf_flags);
        }

        /**
         * Sets the itcLim field value.
         * When bkf.fCol==1, this is the index to limit column of a table column bookmark
         */
        public void SetItcLim(byte value)
        {
            field_2_bkf_flags = (short)itcLim.SetValue(field_2_bkf_flags, value);
        }

        /**
         * When bkf.fCol==1, this is the index to limit column of a table column bookmark
         * @return  the itcLim field value.
         */
        public byte GetItcLim()
        {
            return (byte)itcLim.GetValue(field_2_bkf_flags);
        }

        /**
         * Sets the fCol field value.
         * When 1, this bookmark marks a range of columns in a table specified by (bkf.itcFirst, bkf.itcLim)
         */
        public void SetFCol(bool value)
        {
            field_2_bkf_flags = (short)fCol.SetBoolean(field_2_bkf_flags, value);
        }

        /**
         * When 1, this bookmark marks a range of columns in a table specified by (bkf.itcFirst, bkf.itcLim)
         * @return  the fCol field value.
         */
        public bool IsFCol()
        {
            return fCol.IsSet(field_2_bkf_flags);
        }

    }
}



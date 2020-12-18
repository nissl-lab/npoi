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
using NPOI.Util;
using System.Text;
using System;
namespace NPOI.HWPF.Model.Types
{
    /**
     * List Format Override (LFO).
     * <p>
     * Class and fields descriptions are quoted from Microsoft Office Word 97-2007
     * Binary File Format
     * 
     * NOTE: This source is automatically generated please do not modify this file.
     * Either subclass or remove the record in src/types/defInitions.
     * 
     * @author Sergey Vladimirov; according to Microsoft Office Word 97-2007 Binary
     *         File Format Specification [*.doc]
     */

    public abstract class LFOAbstractType : BaseObject
    {

        protected int field_1_lsid;
        protected int field_2_reserved1;
        protected int field_3_reserved2;
        protected byte field_4_clfolvl;
        protected byte field_5_ibstFltAutoNum;
        protected byte field_6_grfhic;
        private static BitField fHtmlChecked = new BitField(0x01);
        private static BitField fHtmlUnsupported = new BitField(0x02);
        private static BitField fHtmlListTextNotSharpDot = new BitField(0x04);
        private static BitField fHtmlNotPeriod = new BitField(0x08);
        private static BitField fHtmlFirstLineMismatch = new BitField(0x10);
        private static BitField fHtmlTabLeftIndentMismatch = new BitField(
                0x20);
        private static BitField fHtmlHangingIndentBeneathNumber = new BitField(
                0x40);
        private static BitField fHtmlBuiltInBullet = new BitField(0x80);
        protected byte field_7_reserved3;

        protected LFOAbstractType()
        {
        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_lsid = LittleEndian.GetInt(data, 0x0 + offset);
            field_2_reserved1 = LittleEndian.GetInt(data, 0x4 + offset);
            field_3_reserved2 = LittleEndian.GetInt(data, 0x8 + offset);
            field_4_clfolvl = data[0xc + offset];
            field_5_ibstFltAutoNum = data[0xd + offset];
            field_6_grfhic = data[0xe + offset];
            field_7_reserved3 = data[0xf + offset];
        }

        public void Serialize(byte[] data, int offset)
        {
            LittleEndian.PutInt(data, 0x0 + offset, field_1_lsid);
            LittleEndian.PutInt(data, 0x4 + offset, field_2_reserved1);
            LittleEndian.PutInt(data, 0x8 + offset, field_3_reserved2);
            data[0xc + offset] = field_4_clfolvl;
            data[0xd + offset] = field_5_ibstFltAutoNum;
            data[0xe + offset] = field_6_grfhic;
            data[0xf + offset] = field_7_reserved3;
        }

        /**
         * Size of record
         */
        public static int GetSize()
        {
            return 0 + 4 + 4 + 4 + 1 + 1 + 1 + 1;
        }

        public override String ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("[LFO]\n");
            builder.Append("    .lsid                 = ");
            builder.Append(" (").Append(GetLsid()).Append(" )\n");
            builder.Append("    .reserved1            = ");
            builder.Append(" (").Append(GetReserved1()).Append(" )\n");
            builder.Append("    .reserved2            = ");
            builder.Append(" (").Append(GetReserved2()).Append(" )\n");
            builder.Append("    .clfolvl              = ");
            builder.Append(" (").Append(GetClfolvl()).Append(" )\n");
            builder.Append("    .ibstFltAutoNum       = ");
            builder.Append(" (").Append(GetIbstFltAutoNum()).Append(" )\n");
            builder.Append("    .grfhic               = ");
            builder.Append(" (").Append(GetGrfhic()).Append(" )\n");
            builder.Append("         .fHtmlChecked             = ")
                    .Append(IsFHtmlChecked()).Append('\n');
            builder.Append("         .fHtmlUnsupported         = ")
                    .Append(IsFHtmlUnsupported()).Append('\n');
            builder.Append("         .fHtmlListTextNotSharpDot     = ")
                    .Append(IsFHtmlListTextNotSharpDot()).Append('\n');
            builder.Append("         .fHtmlNotPeriod           = ")
                    .Append(IsFHtmlNotPeriod()).Append('\n');
            builder.Append("         .fHtmlFirstLineMismatch     = ")
                    .Append(IsFHtmlFirstLineMismatch()).Append('\n');
            builder.Append("         .fHtmlTabLeftIndentMismatch     = ")
                    .Append(IsFHtmlTabLeftIndentMismatch()).Append('\n');
            builder.Append("         .fHtmlHangingIndentBeneathNumber     = ")
                    .Append(IsFHtmlHangingIndentBeneathNumber()).Append('\n');
            builder.Append("         .fHtmlBuiltInBullet       = ")
                    .Append(IsFHtmlBuiltInBullet()).Append('\n');
            builder.Append("    .reserved3            = ");
            builder.Append(" (").Append(GetReserved3()).Append(" )\n");

            builder.Append("[/LFO]\n");
            return builder.ToString();
        }

        /**
         * List ID of corresponding LSTF (see LSTF).
         */
        public int GetLsid()
        {
            return field_1_lsid;
        }

        /**
         * List ID of corresponding LSTF (see LSTF).
         */
        public void SetLsid(int field_1_lsid)
        {
            this.field_1_lsid = field_1_lsid;
        }

        /**
         * Reserved.
         */
        public int GetReserved1()
        {
            return field_2_reserved1;
        }

        /**
         * Reserved.
         */
        public void SetReserved1(int field_2_reserved1)
        {
            this.field_2_reserved1 = field_2_reserved1;
        }

        /**
         * Reserved.
         */
        public int GetReserved2()
        {
            return field_3_reserved2;
        }

        /**
         * Reserved.
         */
        public void SetReserved2(int field_3_reserved2)
        {
            this.field_3_reserved2 = field_3_reserved2;
        }

        /**
         * Count of levels whose format is overridden (see LFOLVL).
         */
        public byte GetClfolvl()
        {
            return field_4_clfolvl;
        }

        /**
         * Count of levels whose format is overridden (see LFOLVL).
         */
        public void SetClfolvl(byte field_4_clfolvl)
        {
            this.field_4_clfolvl = field_4_clfolvl;
        }

        /**
         * Used for AUTONUM field emulation.
         */
        public byte GetIbstFltAutoNum()
        {
            return field_5_ibstFltAutoNum;
        }

        /**
         * Used for AUTONUM field emulation.
         */
        public void SetIbstFltAutoNum(byte field_5_ibstFltAutoNum)
        {
            this.field_5_ibstFltAutoNum = field_5_ibstFltAutoNum;
        }

        /**
         * HTML compatibility flags.
         */
        public byte GetGrfhic()
        {
            return field_6_grfhic;
        }

        /**
         * HTML compatibility flags.
         */
        public void SetGrfhic(byte field_6_grfhic)
        {
            this.field_6_grfhic = field_6_grfhic;
        }

        /**
         * Reserved.
         */
        public byte GetReserved3()
        {
            return field_7_reserved3;
        }

        /**
         * Reserved.
         */
        public void SetReserved3(byte field_7_reserved3)
        {
            this.field_7_reserved3 = field_7_reserved3;
        }

        /**
         * Sets the fHtmlChecked field value. Checked
         */
        public void SetFHtmlChecked(bool value)
        {
            field_6_grfhic = (byte)fHtmlChecked.SetBoolean(field_6_grfhic, value);
        }

        /**
         * Checked
         * 
         * @return the fHtmlChecked field value.
         */
        public bool IsFHtmlChecked()
        {
            return fHtmlChecked.IsSet(field_6_grfhic);
        }

        /**
         * Sets the fHtmlUnsupported field value. The numbering sequence or format
         * is unsupported (includes tab & size)
         */
        public void SetFHtmlUnsupported(bool value)
        {
            field_6_grfhic = (byte)fHtmlUnsupported.SetBoolean(field_6_grfhic,
                    value);
        }

        /**
         * The numbering sequence or format is unsupported (includes tab & size)
         * 
         * @return the fHtmlUnsupported field value.
         */
        public bool IsFHtmlUnsupported()
        {
            return fHtmlUnsupported.IsSet(field_6_grfhic);
        }

        /**
         * Sets the fHtmlListTextNotSharpDot field value. The list text is not "#."
         */
        public void SetFHtmlListTextNotSharpDot(bool value)
        {
            field_6_grfhic = (byte)fHtmlListTextNotSharpDot.SetBoolean(
                    field_6_grfhic, value);
        }

        /**
         * The list text is not "#."
         * 
         * @return the fHtmlListTextNotSharpDot field value.
         */
        public bool IsFHtmlListTextNotSharpDot()
        {
            return fHtmlListTextNotSharpDot.IsSet(field_6_grfhic);
        }

        /**
         * Sets the fHtmlNotPeriod field value. Something other than a period is
         * used
         */
        public void SetFHtmlNotPeriod(bool value)
        {
            field_6_grfhic = (byte)fHtmlNotPeriod.SetBoolean(field_6_grfhic,
                    value);
        }

        /**
         * Something other than a period is used
         * 
         * @return the fHtmlNotPeriod field value.
         */
        public bool IsFHtmlNotPeriod()
        {
            return fHtmlNotPeriod.IsSet(field_6_grfhic);
        }

        /**
         * Sets the fHtmlFirstLineMismatch field value. First line indent mismatch
         */
        public void SetFHtmlFirstLineMismatch(bool value)
        {
            field_6_grfhic = (byte)fHtmlFirstLineMismatch.SetBoolean(
                    field_6_grfhic, value);
        }

        /**
         * First line indent mismatch
         * 
         * @return the fHtmlFirstLineMismatch field value.
         */
        public bool IsFHtmlFirstLineMismatch()
        {
            return fHtmlFirstLineMismatch.IsSet(field_6_grfhic);
        }

        /**
         * Sets the fHtmlTabLeftIndentMismatch field value. The list tab and the
         * dxaLeft don't match (need table?)
         */
        public void SetFHtmlTabLeftIndentMismatch(bool value)
        {
            field_6_grfhic = (byte)fHtmlTabLeftIndentMismatch.SetBoolean(
                    field_6_grfhic, value);
        }

        /**
         * The list tab and the dxaLeft don't match (need table?)
         * 
         * @return the fHtmlTabLeftIndentMismatch field value.
         */
        public bool IsFHtmlTabLeftIndentMismatch()
        {
            return fHtmlTabLeftIndentMismatch.IsSet(field_6_grfhic);
        }

        /**
         * Sets the fHtmlHangingIndentBeneathNumber field value. The hanging indent
         * falls beneath the number (need plain text)
         */
        public void SetFHtmlHangingIndentBeneathNumber(bool value)
        {
            field_6_grfhic = (byte)fHtmlHangingIndentBeneathNumber.SetBoolean(
                    field_6_grfhic, value);
        }

        /**
         * The hanging indent falls beneath the number (need plain text)
         * 
         * @return the fHtmlHangingIndentBeneathNumber field value.
         */
        public bool IsFHtmlHangingIndentBeneathNumber()
        {
            return fHtmlHangingIndentBeneathNumber.IsSet(field_6_grfhic);
        }

        /**
         * Sets the fHtmlBuiltInBullet field value. A built-in HTML bullet
         */
        public void SetFHtmlBuiltInBullet(bool value)
        {
            field_6_grfhic = (byte)fHtmlBuiltInBullet.SetBoolean(field_6_grfhic,
                    value);
        }

        /**
         * A built-in HTML bullet
         * 
         * @return the fHtmlBuiltInBullet field value.
         */
        public bool IsFHtmlBuiltInBullet()
        {
            return fHtmlBuiltInBullet.IsSet(field_6_grfhic);
        }

    }
}



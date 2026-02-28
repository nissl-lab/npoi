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
     * File Shape Address (FSPA).
     * <p>
     * Class and fields descriptions are quoted from Microsoft Office Word 97-2007
     * Binary File Format
     * 
     * <p>
     * NOTE: This source is automatically generated please do not modify this file.
     * Either subclass or remove the record in src/types/defInitions.
     * <p>
     * This class is internal. It content or properties may change without notice
     * due to Changes in our knowledge of internal Microsoft Word binary structures.
     * 
     * @author Sergey Vladimirov; according to Microsoft Office Word 97-2007 Binary
     *         File Format Specification [*.doc]
     */

    public abstract class FSPAAbstractType : BaseObject
    {

        protected int field_1_spid;
        protected int field_2_xaLeft;
        protected int field_3_yaTop;
        protected int field_4_xaRight;
        protected int field_5_yaBottom;
        protected short field_6_flags;
        private static BitField fHdr = new BitField(0x0001);
        private static BitField bx = new BitField(0x0006);
        private static BitField by = new BitField(0x0018);
        private static BitField wr = new BitField(0x01E0);
        private static BitField wrk = new BitField(0x1E00);
        private static BitField fRcaSimple = new BitField(0x2000);
        private static BitField fBelowText = new BitField(0x4000);
        private static BitField fAnchorLock = new BitField(0x8000);
        protected int field_7_cTxbx;

        protected FSPAAbstractType()
        {
        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_spid = LittleEndian.GetInt(data, 0x0 + offset);
            field_2_xaLeft = LittleEndian.GetInt(data, 0x4 + offset);
            field_3_yaTop = LittleEndian.GetInt(data, 0x8 + offset);
            field_4_xaRight = LittleEndian.GetInt(data, 0xc + offset);
            field_5_yaBottom = LittleEndian.GetInt(data, 0x10 + offset);
            field_6_flags = LittleEndian.GetShort(data, 0x14 + offset);
            field_7_cTxbx = LittleEndian.GetInt(data, 0x16 + offset);
        }

        public void Serialize(byte[] data, int offset)
        {
            LittleEndian.PutInt(data, 0x0 + offset, field_1_spid);
            LittleEndian.PutInt(data, 0x4 + offset, field_2_xaLeft);
            LittleEndian.PutInt(data, 0x8 + offset, field_3_yaTop);
            LittleEndian.PutInt(data, 0xc + offset, field_4_xaRight);
            LittleEndian.PutInt(data, 0x10 + offset, field_5_yaBottom);
            LittleEndian.PutShort(data, 0x14 + offset, (short)field_6_flags);
            LittleEndian.PutInt(data, 0x16 + offset, field_7_cTxbx);
        }

        /**
         * Size of record
         */
        public static int GetSize()
        {
            return 0 + 4 + 4 + 4 + 4 + 4 + 2 + 4;
        }

        public override String ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("[FSPA]\n");
            builder.Append("    .spid                 = ");
            builder.Append(" (").Append(GetSpid()).Append(" )\n");
            builder.Append("    .xaLeft               = ");
            builder.Append(" (").Append(GetXaLeft()).Append(" )\n");
            builder.Append("    .yaTop                = ");
            builder.Append(" (").Append(GetYaTop()).Append(" )\n");
            builder.Append("    .xaRight              = ");
            builder.Append(" (").Append(GetXaRight()).Append(" )\n");
            builder.Append("    .yaBottom             = ");
            builder.Append(" (").Append(GetYaBottom()).Append(" )\n");
            builder.Append("    .flags                = ");
            builder.Append(" (").Append(GetFlags()).Append(" )\n");
            builder.Append("         .fHdr                     = ").Append(IsFHdr()).Append('\n');
            builder.Append("         .bx                       = ").Append(GetBx()).Append('\n');
            builder.Append("         .by                       = ").Append(GetBy()).Append('\n');
            builder.Append("         .wr                       = ").Append(GetWr()).Append('\n');
            builder.Append("         .wrk                      = ").Append(GetWrk()).Append('\n');
            builder.Append("         .fRcaSimple               = ").Append(IsFRcaSimple()).Append('\n');
            builder.Append("         .fBelowText               = ").Append(IsFBelowText()).Append('\n');
            builder.Append("         .fAnchorLock              = ").Append(IsFAnchorLock()).Append('\n');
            builder.Append("    .cTxbx                = ");
            builder.Append(" (").Append(GetCTxbx()).Append(" )\n");

            builder.Append("[/FSPA]\n");
            return builder.ToString();
        }

        /**
         * Shape Identifier. Used in conjunction with the office art data (found via fcDggInfo in the FIB) to find the actual data for this shape.
         */

        public int GetSpid()
        {
            return field_1_spid;
        }

        /**
         * Shape Identifier. Used in conjunction with the office art data (found via fcDggInfo in the FIB) to find the actual data for this shape.
         */

        public void SetSpid(int field_1_spid)
        {
            this.field_1_spid = field_1_spid;
        }

        /**
         * Left of rectangle enclosing shape relative to the origin of the shape.
         */

        public int GetXaLeft()
        {
            return field_2_xaLeft;
        }

        /**
         * Left of rectangle enclosing shape relative to the origin of the shape.
         */

        public void SetXaLeft(int field_2_xaLeft)
        {
            this.field_2_xaLeft = field_2_xaLeft;
        }

        /**
         * Top of rectangle enclosing shape relative to the origin of the shape.
         */

        public int GetYaTop()
        {
            return field_3_yaTop;
        }

        /**
         * Top of rectangle enclosing shape relative to the origin of the shape.
         */

        public void SetYaTop(int field_3_yaTop)
        {
            this.field_3_yaTop = field_3_yaTop;
        }

        /**
         * Right of rectangle enclosing shape relative to the origin of the shape.
         */

        public int GetXaRight()
        {
            return field_4_xaRight;
        }

        /**
         * Right of rectangle enclosing shape relative to the origin of the shape.
         */

        public void SetXaRight(int field_4_xaRight)
        {
            this.field_4_xaRight = field_4_xaRight;
        }

        /**
         * Bottom of the rectangle enclosing shape relative to the origin of the shape.
         */

        public int GetYaBottom()
        {
            return field_5_yaBottom;
        }

        /**
         * Bottom of the rectangle enclosing shape relative to the origin of the shape.
         */

        public void SetYaBottom(int field_5_yaBottom)
        {
            this.field_5_yaBottom = field_5_yaBottom;
        }

        /**
         * Get the flags field for the FSPA record.
         */

        public short GetFlags()
        {
            return field_6_flags;
        }

        /**
         * Set the flags field for the FSPA record.
         */

        public void SetFlags(short field_6_flags)
        {
            this.field_6_flags = field_6_flags;
        }

        /**
         * Count of textboxes in shape (undo doc only).
         */

        public int GetCTxbx()
        {
            return field_7_cTxbx;
        }

        /**
         * Count of textboxes in shape (undo doc only).
         */

        public void SetCTxbx(int field_7_cTxbx)
        {
            this.field_7_cTxbx = field_7_cTxbx;
        }

        /**
         * Sets the fHdr field value.
         * 1 in the undo doc when shape is from the header doc, 0 otherwise (undefined when not in the undo doc)
         */

        public void SetFHdr(bool value)
        {
            field_6_flags = (short)fHdr.SetBoolean(field_6_flags, value);
        }

        /**
         * 1 in the undo doc when shape is from the header doc, 0 otherwise (undefined when not in the undo doc)
         * @return  the fHdr field value.
         */

        public bool IsFHdr()
        {
            return fHdr.IsSet(field_6_flags);
        }

        /**
         * Sets the bx field value.
         * X position of shape relative to anchor CP
         */

        public void SetBx(byte value)
        {
            field_6_flags = (short)bx.SetValue(field_6_flags, value);
        }

        /**
         * X position of shape relative to anchor CP
         * @return  the bx field value.
         */

        public byte GetBx()
        {
            return (byte)bx.GetValue(field_6_flags);
        }

        /**
         * Sets the by field value.
         * Y position of shape relative to anchor CP
         */

        public void SetBy(byte value)
        {
            field_6_flags = (short)by.SetValue(field_6_flags, value);
        }

        /**
         * Y position of shape relative to anchor CP
         * @return  the by field value.
         */

        public byte GetBy()
        {
            return (byte)by.GetValue(field_6_flags);
        }

        /**
         * Sets the wr field value.
         * Text wrapping mode
         */

        public void SetWr(byte value)
        {
            field_6_flags = (short)wr.SetValue(field_6_flags, value);
        }

        /**
         * Text wrapping mode
         * @return  the wr field value.
         */

        public byte GetWr()
        {
            return (byte)wr.GetValue(field_6_flags);
        }

        /**
         * Sets the wrk field value.
         * Text wrapping mode type (valid only for wrapping modes 2 and 4
         */

        public void SetWrk(byte value)
        {
            field_6_flags = (short)wrk.SetValue(field_6_flags, value);
        }

        /**
         * Text wrapping mode type (valid only for wrapping modes 2 and 4
         * @return  the wrk field value.
         */

        public byte GetWrk()
        {
            return (byte)wrk.GetValue(field_6_flags);
        }

        /**
         * Sets the fRcaSimple field value.
         * When Set, temporarily overrides bx, by, forcing the xaLeft, xaRight, yaTop, and yaBottom fields to all be page relative.
         */

        public void SetFRcaSimple(bool value)
        {
            field_6_flags = (short)fRcaSimple.SetBoolean(field_6_flags, value);
        }

        /**
         * When Set, temporarily overrides bx, by, forcing the xaLeft, xaRight, yaTop, and yaBottom fields to all be page relative.
         * @return  the fRcaSimple field value.
         */

        public bool IsFRcaSimple()
        {
            return fRcaSimple.IsSet(field_6_flags);
        }

        /**
         * Sets the fBelowText field value.
         * 
         */

        public void SetFBelowText(bool value)
        {
            field_6_flags = (short)fBelowText.SetBoolean(field_6_flags, value);
        }

        /**
         * 
         * @return  the fBelowText field value.
         */

        public bool IsFBelowText()
        {
            return fBelowText.IsSet(field_6_flags);
        }

        /**
         * Sets the fAnchorLock field value.
         * 
         */

        public void SetFAnchorLock(bool value)
        {
            field_6_flags = (short)fAnchorLock.SetBoolean(field_6_flags, value);
        }

        /**
         * 
         * @return  the fAnchorLock field value.
         */

        public bool IsFAnchorLock()
        {
            return fAnchorLock.IsSet(field_6_flags);
        }

    }
}



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
     * Table Cell Descriptor.
     * NOTE: This source Is automatically generated please do not modify this file.  Either subclass or
     *       remove the record in src/records/definitions.

     * @author S. Ryan Ackley
     */
    public abstract class TCAbstractType : BaseObject
    {

        protected short field_1_rgf;
        private static BitField fFirstMerged = BitFieldFactory.GetInstance(0x0001);
        private static BitField fMerged = BitFieldFactory.GetInstance(0x0002);
        private static BitField fVertical = BitFieldFactory.GetInstance(0x0004);
        private static BitField fBackward = BitFieldFactory.GetInstance(0x0008);
        private static BitField fRotateFont = BitFieldFactory.GetInstance(0x0010);
        private static BitField fVertMerge = BitFieldFactory.GetInstance(0x0020);
        private static BitField fVertRestart = BitFieldFactory.GetInstance(0x0040);
        private static BitField vertAlign = BitFieldFactory.GetInstance(0x0180);
        private static BitField ftsWidth = new BitField(0x0E00);
        private static BitField fFitText = new BitField(0x1000);
        private static BitField fNoWrap = new BitField(0x2000);
        private static BitField fUnused = new BitField(0xC000);
        protected short field_2_wWidth;
        protected short field_3_wCellPaddingLeft;
        protected short field_4_wCellPaddingTop;
        protected short field_5_wCellPaddingBottom;
        protected short field_6_wCellPaddingRight;
        protected byte field_7_ftsCellPaddingLeft;
        protected byte field_8_ftsCellPaddingTop;
        protected byte field_9_ftsCellPaddingBottom;
        protected byte field_10_ftsCellPaddingRight;
        protected short field_11_wCellSpacingLeft;
        protected short field_12_wCellSpacingTop;
        protected short field_13_wCellSpacingBottom;
        protected short field_14_wCellSpacingRight;
        protected byte field_15_ftsCellSpacingLeft;
        protected byte field_16_ftsCellSpacingTop;
        protected byte field_17_ftsCellSpacingBottom;
        protected byte field_18_ftsCellSpacingRight;
        protected BorderCode field_19_brcTop;
        protected BorderCode field_20_brcLeft;
        protected BorderCode field_21_brcBottom;
        protected BorderCode field_22_brcRight;


        public TCAbstractType()
        {

        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[TC]\n");

            buffer.Append("    .rgf                  = ");
            buffer.Append(" (").Append(GetRgf()).Append(" )\n");
            buffer.Append("         .fFirstMerged             = ").Append(IsFFirstMerged()).Append('\n');
            buffer.Append("         .fMerged                  = ").Append(IsFMerged()).Append('\n');
            buffer.Append("         .fVertical                = ").Append(IsFVertical()).Append('\n');
            buffer.Append("         .fBackward                = ").Append(IsFBackward()).Append('\n');
            buffer.Append("         .fRotateFont              = ").Append(IsFRotateFont()).Append('\n');
            buffer.Append("         .fVertMerge               = ").Append(IsFVertMerge()).Append('\n');
            buffer.Append("         .fVertRestart             = ").Append(IsFVertRestart()).Append('\n');
            buffer.Append("         .vertAlign                = ").Append(GetVertAlign()).Append('\n');
            buffer.Append("         .ftsWidth                 = ").Append(GetFtsWidth()).Append('\n');
            buffer.Append("         .fFitText                 = ").Append(IsFFitText()).Append('\n');
            buffer.Append("         .fNoWrap                  = ").Append(IsFNoWrap()).Append('\n');
            buffer.Append("         .fUnused                  = ").Append(GetFUnused()).Append('\n');

            buffer.Append("    .wWidth               = ");
            buffer.Append(" (").Append(GetWWidth()).Append(" )\n");

            buffer.Append("    .wCellPaddingLeft     = ");
            buffer.Append(" (").Append(GetWCellPaddingLeft()).Append(" )\n");

            buffer.Append("    .wCellPaddingTop      = ");
            buffer.Append(" (").Append(GetWCellPaddingTop()).Append(" )\n");

            buffer.Append("    .wCellPaddingBottom   = ");
            buffer.Append(" (").Append(GetWCellPaddingBottom()).Append(" )\n");

            buffer.Append("    .wCellPaddingRight    = ");
            buffer.Append(" (").Append(GetWCellPaddingRight()).Append(" )\n");

            buffer.Append("    .ftsCellPaddingLeft   = ");
            buffer.Append(" (").Append(GetFtsCellPaddingLeft()).Append(" )\n");

            buffer.Append("    .ftsCellPaddingTop    = ");
            buffer.Append(" (").Append(GetFtsCellPaddingTop()).Append(" )\n");

            buffer.Append("    .ftsCellPaddingBottom = ");
            buffer.Append(" (").Append(GetFtsCellPaddingBottom()).Append(" )\n");

            buffer.Append("    .ftsCellPaddingRight  = ");
            buffer.Append(" (").Append(GetFtsCellPaddingRight()).Append(" )\n");

            buffer.Append("    .wCellSpacingLeft     = ");
            buffer.Append(" (").Append(GetWCellSpacingLeft()).Append(" )\n");

            buffer.Append("    .wCellSpacingTop      = ");
            buffer.Append(" (").Append(GetWCellSpacingTop()).Append(" )\n");

            buffer.Append("    .wCellSpacingBottom   = ");
            buffer.Append(" (").Append(GetWCellSpacingBottom()).Append(" )\n");

            buffer.Append("    .wCellSpacingRight    = ");
            buffer.Append(" (").Append(GetWCellSpacingRight()).Append(" )\n");

            buffer.Append("    .ftsCellSpacingLeft   = ");
            buffer.Append(" (").Append(GetFtsCellSpacingLeft()).Append(" )\n");

            buffer.Append("    .ftsCellSpacingTop    = ");
            buffer.Append(" (").Append(GetFtsCellSpacingTop()).Append(" )\n");

            buffer.Append("    .ftsCellSpacingBottom = ");
            buffer.Append(" (").Append(GetFtsCellSpacingBottom()).Append(" )\n");

            buffer.Append("    .ftsCellSpacingRight  = ");
            buffer.Append(" (").Append(GetFtsCellSpacingRight()).Append(" )\n");

            buffer.Append("    .brcTop               = ");
            buffer.Append(" (").Append(GetBrcTop()).Append(" )\n");

            buffer.Append("    .brcLeft              = ");
            buffer.Append(" (").Append(GetBrcLeft()).Append(" )\n");

            buffer.Append("    .brcBottom            = ");
            buffer.Append(" (").Append(GetBrcBottom()).Append(" )\n");

            buffer.Append("    .brcRight             = ");
            buffer.Append(" (").Append(GetBrcRight()).Append(" )\n");

            buffer.Append("[/TC]\n");
            return buffer.ToString();
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public int GetSize()
        {
            return 4 + +2 + 2 + 2 + 2 + 2 + 2 + 1 + 1 + 1 + 1 + 2 + 2 + 2 + 2 + 1 + 1 + 1 + 1 + 4 + 4 + 4 + 4;
        }



        /**
         * Get the rgf field for the TC record.
         */
        public short GetRgf()
        {
            return field_1_rgf;
        }

        /**
         * Set the rgf field for the TC record.
         */
        public void SetRgf(short field_1_rgf)
        {
            this.field_1_rgf = field_1_rgf;
        }

        /**
     * Get the wWidth field for the TC record.
     */
        public short GetWWidth()
        {
            return field_2_wWidth;
        }

        /**
         * Set the wWidth field for the TC record.
         */
        public void SetWWidth(short field_2_wWidth)
        {
            this.field_2_wWidth = field_2_wWidth;
        }

        /**
         * Get the wCellPaddingLeft field for the TC record.
         */
        public short GetWCellPaddingLeft()
        {
            return field_3_wCellPaddingLeft;
        }

        /**
         * Set the wCellPaddingLeft field for the TC record.
         */
        public void SetWCellPaddingLeft(short field_3_wCellPaddingLeft)
        {
            this.field_3_wCellPaddingLeft = field_3_wCellPaddingLeft;
        }

        /**
         * Get the wCellPaddingTop field for the TC record.
         */
        public short GetWCellPaddingTop()
        {
            return field_4_wCellPaddingTop;
        }

        /**
         * Set the wCellPaddingTop field for the TC record.
         */
        public void SetWCellPaddingTop(short field_4_wCellPaddingTop)
        {
            this.field_4_wCellPaddingTop = field_4_wCellPaddingTop;
        }

        /**
         * Get the wCellPaddingBottom field for the TC record.
         */
        public short GetWCellPaddingBottom()
        {
            return field_5_wCellPaddingBottom;
        }

        /**
         * Set the wCellPaddingBottom field for the TC record.
         */
        public void SetWCellPaddingBottom(short field_5_wCellPaddingBottom)
        {
            this.field_5_wCellPaddingBottom = field_5_wCellPaddingBottom;
        }

        /**
         * Get the wCellPaddingRight field for the TC record.
         */
        public short GetWCellPaddingRight()
        {
            return field_6_wCellPaddingRight;
        }

        /**
         * Set the wCellPaddingRight field for the TC record.
         */
        public void SetWCellPaddingRight(short field_6_wCellPaddingRight)
        {
            this.field_6_wCellPaddingRight = field_6_wCellPaddingRight;
        }

        /**
         * Get the ftsCellPaddingLeft field for the TC record.
         */
        public byte GetFtsCellPaddingLeft()
        {
            return field_7_ftsCellPaddingLeft;
        }

        /**
         * Set the ftsCellPaddingLeft field for the TC record.
         */
        public void SetFtsCellPaddingLeft(byte field_7_ftsCellPaddingLeft)
        {
            this.field_7_ftsCellPaddingLeft = field_7_ftsCellPaddingLeft;
        }

        /**
         * Get the ftsCellPaddingTop field for the TC record.
         */
        public byte GetFtsCellPaddingTop()
        {
            return field_8_ftsCellPaddingTop;
        }

        /**
         * Set the ftsCellPaddingTop field for the TC record.
         */
        public void SetFtsCellPaddingTop(byte field_8_ftsCellPaddingTop)
        {
            this.field_8_ftsCellPaddingTop = field_8_ftsCellPaddingTop;
        }

        /**
         * Get the ftsCellPaddingBottom field for the TC record.
         */
        public byte GetFtsCellPaddingBottom()
        {
            return field_9_ftsCellPaddingBottom;
        }

        /**
         * Set the ftsCellPaddingBottom field for the TC record.
         */
        public void SetFtsCellPaddingBottom(byte field_9_ftsCellPaddingBottom)
        {
            this.field_9_ftsCellPaddingBottom = field_9_ftsCellPaddingBottom;
        }

        /**
         * Get the ftsCellPaddingRight field for the TC record.
         */
        public byte GetFtsCellPaddingRight()
        {
            return field_10_ftsCellPaddingRight;
        }

        /**
         * Set the ftsCellPaddingRight field for the TC record.
         */
        public void SetFtsCellPaddingRight(byte field_10_ftsCellPaddingRight)
        {
            this.field_10_ftsCellPaddingRight = field_10_ftsCellPaddingRight;
        }

        /**
         * Get the wCellSpacingLeft field for the TC record.
         */
        public short GetWCellSpacingLeft()
        {
            return field_11_wCellSpacingLeft;
        }

        /**
         * Set the wCellSpacingLeft field for the TC record.
         */
        public void SetWCellSpacingLeft(short field_11_wCellSpacingLeft)
        {
            this.field_11_wCellSpacingLeft = field_11_wCellSpacingLeft;
        }

        /**
         * Get the wCellSpacingTop field for the TC record.
         */
        public short GetWCellSpacingTop()
        {
            return field_12_wCellSpacingTop;
        }

        /**
         * Set the wCellSpacingTop field for the TC record.
             */
        public void SetWCellSpacingTop(short field_12_wCellSpacingTop)
        {
            this.field_12_wCellSpacingTop = field_12_wCellSpacingTop;
        }

        /**
         * Get the wCellSpacingBottom field for the TC record.
         */
        public short GetWCellSpacingBottom()
        {
            return field_13_wCellSpacingBottom;
        }

        /**
         * Set the wCellSpacingBottom field for the TC record.
         */
        public void SetWCellSpacingBottom(short field_13_wCellSpacingBottom)
        {
            this.field_13_wCellSpacingBottom = field_13_wCellSpacingBottom;
        }

        /**
         * Get the wCellSpacingRight field for the TC record.
         */
        public short GetWCellSpacingRight()
        {
            return field_14_wCellSpacingRight;
        }

        /**
         * Set the wCellSpacingRight field for the TC record.
         */
        public void SetWCellSpacingRight(short field_14_wCellSpacingRight)
        {
            this.field_14_wCellSpacingRight = field_14_wCellSpacingRight;
        }

        /**
     * Get the ftsCellSpacingLeft field for the TC record.
         */
        public byte GetFtsCellSpacingLeft()
        {
            return field_15_ftsCellSpacingLeft;
        }

        /**
         * Set the ftsCellSpacingLeft field for the TC record.
         */
        public void SetFtsCellSpacingLeft(byte field_15_ftsCellSpacingLeft)
        {
            this.field_15_ftsCellSpacingLeft = field_15_ftsCellSpacingLeft;
        }

        /**
         * Get the ftsCellSpacingTop field for the TC record.
         */
        public byte GetFtsCellSpacingTop()
        {
            return field_16_ftsCellSpacingTop;
        }

        /**
         * Set the ftsCellSpacingTop field for the TC record.
         */
        public void SetFtsCellSpacingTop(byte field_16_ftsCellSpacingTop)
        {
            this.field_16_ftsCellSpacingTop = field_16_ftsCellSpacingTop;
        }

        /**
         * Get the ftsCellSpacingBottom field for the TC record.
         */
        public byte GetFtsCellSpacingBottom()
        {
            return field_17_ftsCellSpacingBottom;
        }

        /**
         * Set the ftsCellSpacingBottom field for the TC record.
         */
        public void SetFtsCellSpacingBottom(byte field_17_ftsCellSpacingBottom)
        {
            this.field_17_ftsCellSpacingBottom = field_17_ftsCellSpacingBottom;
        }

        /**
         * Get the ftsCellSpacingRight field for the TC record.
         */
        public byte GetFtsCellSpacingRight()
        {
            return field_18_ftsCellSpacingRight;
        }

        /**
         * Set the ftsCellSpacingRight field for the TC record.
         */
        public void SetFtsCellSpacingRight(byte field_18_ftsCellSpacingRight)
        {
            this.field_18_ftsCellSpacingRight = field_18_ftsCellSpacingRight;
        }

        /**
         * Get the brcTop field for the TC record.
         */
        public BorderCode GetBrcTop()
        {
            return field_19_brcTop;
        }

        /**
         * Set the brcTop field for the TC record.
         */
        public void SetBrcTop(BorderCode field_19_brcTop)
        {
            this.field_19_brcTop = field_19_brcTop;
        }

        /**
         * Get the brcLeft field for the TC record.
         */
        public BorderCode GetBrcLeft()
        {
            return field_20_brcLeft;
        }

        /**
         * Set the brcLeft field for the TC record.
         */
        public void SetBrcLeft(BorderCode field_20_brcLeft)
        {
            this.field_20_brcLeft = field_20_brcLeft;
        }

        /**
         * Get the brcBottom field for the TC record.
         */
        public BorderCode GetBrcBottom()
        {
            return field_21_brcBottom;
        }

        /**
         * Set the brcBottom field for the TC record.
         */
        public void SetBrcBottom(BorderCode field_21_brcBottom)
        {
            this.field_21_brcBottom = field_21_brcBottom;
        }

        /**
         * Get the brcRight field for the TC record.
         */
        public BorderCode GetBrcRight()
        {
            return field_22_brcRight;
        }

        /**
         * Set the brcRight field for the TC record.
         */
        public void SetBrcRight(BorderCode field_22_brcRight)
        {
            this.field_22_brcRight = field_22_brcRight;
        }

        /**
         * Sets the fFirstMerged field value.
         * 
         */
        public void SetFFirstMerged(bool value)
        {
            field_1_rgf = (short)fFirstMerged.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fFirstMerged field value.
         */
        public bool IsFFirstMerged()
        {
            return fFirstMerged.IsSet(field_1_rgf);

        }

        /**
         * Sets the fMerged field value.
         * 
         */
        public void SetFMerged(bool value)
        {
            field_1_rgf = (short)fMerged.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fMerged field value.
         */
        public bool IsFMerged()
        {
            return fMerged.IsSet(field_1_rgf);

        }

        /**
         * Sets the fVertical field value.
         * 
         */
        public void SetFVertical(bool value)
        {
            field_1_rgf = (short)fVertical.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fVertical field value.
         */
        public bool IsFVertical()
        {
            return fVertical.IsSet(field_1_rgf);

        }

        /**
         * Sets the fBackward field value.
         * 
         */
        public void SetFBackward(bool value)
        {
            field_1_rgf = (short)fBackward.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fBackward field value.
         */
        public bool IsFBackward()
        {
            return fBackward.IsSet(field_1_rgf);

        }

        /**
         * Sets the fRotateFont field value.
         * 
         */
        public void SetFRotateFont(bool value)
        {
            field_1_rgf = (short)fRotateFont.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fRotateFont field value.
         */
        public bool IsFRotateFont()
        {
            return fRotateFont.IsSet(field_1_rgf);

        }

        /**
         * Sets the fVertMerge field value.
         * 
         */
        public void SetFVertMerge(bool value)
        {
            field_1_rgf = (short)fVertMerge.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fVertMerge field value.
         */
        public bool IsFVertMerge()
        {
            return fVertMerge.IsSet(field_1_rgf);

        }

        /**
         * Sets the fVertRestart field value.
         * 
         */
        public void SetFVertRestart(bool value)
        {
            field_1_rgf = (short)fVertRestart.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fVertRestart field value.
         */
        public bool IsFVertRestart()
        {
            return fVertRestart.IsSet(field_1_rgf);

        }

        /**
         * Sets the vertAlign field value.
         * 
         */
        public void SetVertAlign(byte value)
        {
            field_1_rgf = (short)vertAlign.SetValue(field_1_rgf, value);


        }

        /**
         * 
         * @return  the vertAlign field value.
         */
        public byte GetVertAlign()
        {
            return (byte)vertAlign.GetValue(field_1_rgf);

        }

        /**
         * Sets the ftsWidth field value.
         * 
         */
        public void SetFtsWidth(byte value)
        {
            field_1_rgf = (short)ftsWidth.SetValue(field_1_rgf, value);


        }

        /**
         * 
         * @return  the ftsWidth field value.
         */
        public byte GetFtsWidth()
        {
            return (byte)ftsWidth.GetValue(field_1_rgf);

        }

        /**
         * Sets the fFitText field value.
         * 
         */
        public void SetFFitText(bool value)
        {
            field_1_rgf = (short)fFitText.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fFitText field value.
         */
        public bool IsFFitText()
        {
            return fFitText.IsSet(field_1_rgf);

        }

        /**
         * Sets the fNoWrap field value.
         * 
         */
        public void SetFNoWrap(bool value)
        {
            field_1_rgf = (short)fNoWrap.SetBoolean(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fNoWrap field value.
         */
        public bool IsFNoWrap()
        {
            return fNoWrap.IsSet(field_1_rgf);

        }

        /**
         * Sets the fUnused field value.
         * 
         */
        public void SetFUnused(byte value)
        {
            field_1_rgf = (short)fUnused.SetValue(field_1_rgf, value);


        }

        /**
         * 
         * @return  the fUnused field value.
         */
        public byte GetFUnused()
        {
            return (byte)fUnused.GetValue(field_1_rgf);

        }


    }
}





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
        protected short field_2_unused;
        protected BorderCode field_3_brcTop;
        protected BorderCode field_4_brcLeft;
        protected BorderCode field_5_brcBottom;
        protected BorderCode field_6_brcRight;


        public TCAbstractType()
        {

        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_rgf = LittleEndian.GetShort(data, 0x0 + offset);
            field_2_unused = LittleEndian.GetShort(data, 0x2 + offset);
            field_3_brcTop = new BorderCode(data, 0x4 + offset);
            field_4_brcLeft = new BorderCode(data, 0x8 + offset);
            field_5_brcBottom = new BorderCode(data, 0xc + offset);
            field_6_brcRight = new BorderCode(data, 0x10 + offset);

        }

        public void Serialize(byte[] data, int offset)
        {
            LittleEndian.PutShort(data, 0x0 + offset, (short)field_1_rgf); ;
            LittleEndian.PutShort(data, 0x2 + offset, (short)field_2_unused); ;
            field_3_brcTop.Serialize(data, 0x4 + offset); ;
            field_4_brcLeft.Serialize(data, 0x8 + offset); ;
            field_5_brcBottom.Serialize(data, 0xc + offset); ;
            field_6_brcRight.Serialize(data, 0x10 + offset); ;

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

            buffer.Append("    .unused               = ");
            buffer.Append(" (").Append(GetUnused()).Append(" )\n");

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
            return 4 + +2 + 2 + 4 + 4 + 4 + 4;
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
         * Get the unused field for the TC record.
         */
        public short GetUnused()
        {
            return field_2_unused;
        }

        /**
         * Set the unused field for the TC record.
         */
        public void SetUnused(short field_2_unused)
        {
            this.field_2_unused = field_2_unused;
        }

        /**
         * Get the brcTop field for the TC record.
         */
        public BorderCode GetBrcTop()
        {
            return field_3_brcTop;
        }

        /**
         * Set the brcTop field for the TC record.
         */
        public void SetBrcTop(BorderCode field_3_brcTop)
        {
            this.field_3_brcTop = field_3_brcTop;
        }

        /**
         * Get the brcLeft field for the TC record.
         */
        public BorderCode GetBrcLeft()
        {
            return field_4_brcLeft;
        }

        /**
         * Set the brcLeft field for the TC record.
         */
        public void SetBrcLeft(BorderCode field_4_brcLeft)
        {
            this.field_4_brcLeft = field_4_brcLeft;
        }

        /**
         * Get the brcBottom field for the TC record.
         */
        public BorderCode GetBrcBottom()
        {
            return field_5_brcBottom;
        }

        /**
         * Set the brcBottom field for the TC record.
         */
        public void SetBrcBottom(BorderCode field_5_brcBottom)
        {
            this.field_5_brcBottom = field_5_brcBottom;
        }

        /**
         * Get the brcRight field for the TC record.
         */
        public BorderCode GetBrcRight()
        {
            return field_6_brcRight;
        }

        /**
         * Set the brcRight field for the TC record.
         */
        public void SetBrcRight(BorderCode field_6_brcRight)
        {
            this.field_6_brcRight = field_6_brcRight;
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


    }
}
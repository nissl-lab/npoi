
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
     * Table Properties.
     * NOTE: This source Is automatically generated please do not modify this file.  Either subclass or
     *       remove the record in src/records/definitions.

     * @author S. Ryan Ackley
     */
    public abstract class TAPAbstractType:BaseObject 
    {

        protected int field_1_jc;
        protected int field_2_dxaGapHalf;
        protected int field_3_dyaRowHeight;
        protected bool field_4_fCantSplit;
        protected bool field_5_fTableHeader;
        protected int field_6_tlp;
        protected short field_7_itcMac;
        protected short[] field_8_rgdxaCenter;
        protected TableCellDescriptor[] field_9_rgtc;
        protected ShadingDescriptor[] field_10_rgshd;
        protected BorderCode field_11_brcBottom;
        protected BorderCode field_12_brcTop;
        protected BorderCode field_13_brcLeft;
        protected BorderCode field_14_brcRight;
        protected BorderCode field_15_brcVertical;
        protected BorderCode field_16_brcHorizontal;


        public TAPAbstractType()
        {

        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public int GetSize()
        {
            return 4 + +2 + 4 + 4 + 0 + 0 + 4 + 2 + 130 + 0 + 0 + 4 + 4 + 4 + 4 + 4 + 4;
        }



        /**
         * Get the jc field for the TAP record.
         */
        public int GetJc()
        {
            return field_1_jc;
        }

        /**
         * Set the jc field for the TAP record.
         */
        public void SetJc(int field_1_jc)
        {
            this.field_1_jc = field_1_jc;
        }

        /**
         * Get the dxaGapHalf field for the TAP record.
         */
        public int GetDxaGapHalf()
        {
            return field_2_dxaGapHalf;
        }

        /**
         * Set the dxaGapHalf field for the TAP record.
         */
        public void SetDxaGapHalf(int field_2_dxaGapHalf)
        {
            this.field_2_dxaGapHalf = field_2_dxaGapHalf;
        }

        /**
         * Get the dyaRowHeight field for the TAP record.
         */
        public int GetDyaRowHeight()
        {
            return field_3_dyaRowHeight;
        }

        /**
         * Set the dyaRowHeight field for the TAP record.
         */
        public void SetDyaRowHeight(int field_3_dyaRowHeight)
        {
            this.field_3_dyaRowHeight = field_3_dyaRowHeight;
        }

        /**
         * Get the fCantSplit field for the TAP record.
         */
        public bool GetFCantSplit()
        {
            return field_4_fCantSplit;
        }

        /**
         * Set the fCantSplit field for the TAP record.
         */
        public void SetFCantSplit(bool field_4_fCantSplit)
        {
            this.field_4_fCantSplit = field_4_fCantSplit;
        }

        /**
         * Get the fTableHeader field for the TAP record.
         */
        public bool GetFTableHeader()
        {
            return field_5_fTableHeader;
        }

        /**
         * Set the fTableHeader field for the TAP record.
         */
        public void SetFTableHeader(bool field_5_fTableHeader)
        {
            this.field_5_fTableHeader = field_5_fTableHeader;
        }

        /**
         * Get the tlp field for the TAP record.
         */
        public int GetTlp()
        {
            return field_6_tlp;
        }

        /**
         * Set the tlp field for the TAP record.
         */
        public void SetTlp(int field_6_tlp)
        {
            this.field_6_tlp = field_6_tlp;
        }

        /**
         * Get the itcMac field for the TAP record.
         */
        public short GetItcMac()
        {
            return field_7_itcMac;
        }

        /**
         * Set the itcMac field for the TAP record.
         */
        public void SetItcMac(short field_7_itcMac)
        {
            this.field_7_itcMac = field_7_itcMac;
        }

        /**
         * Get the rgdxaCenter field for the TAP record.
         */
        public short[] GetRgdxaCenter()
        {
            return field_8_rgdxaCenter;
        }

        /**
         * Set the rgdxaCenter field for the TAP record.
         */
        public void SetRgdxaCenter(short[] field_8_rgdxaCenter)
        {
            this.field_8_rgdxaCenter = field_8_rgdxaCenter;
        }

        /**
         * Get the rgtc field for the TAP record.
         */
        public TableCellDescriptor[] GetRgtc()
        {
            return field_9_rgtc;
        }

        /**
         * Set the rgtc field for the TAP record.
         */
        public void SetRgtc(TableCellDescriptor[] field_9_rgtc)
        {
            this.field_9_rgtc = field_9_rgtc;
        }

        /**
         * Get the rgshd field for the TAP record.
         */
        public ShadingDescriptor[] GetRgshd()
        {
            return field_10_rgshd;
        }

        /**
         * Set the rgshd field for the TAP record.
         */
        public void SetRgshd(ShadingDescriptor[] field_10_rgshd)
        {
            this.field_10_rgshd = field_10_rgshd;
        }

        /**
         * Get the brcBottom field for the TAP record.
         */
        public BorderCode GetBrcBottom()
        {
            return field_11_brcBottom;
        }

        /**
         * Set the brcBottom field for the TAP record.
         */
        public void SetBrcBottom(BorderCode field_11_brcBottom)
        {
            this.field_11_brcBottom = field_11_brcBottom;
        }

        /**
         * Get the brcTop field for the TAP record.
         */
        public BorderCode GetBrcTop()
        {
            return field_12_brcTop;
        }

        /**
         * Set the brcTop field for the TAP record.
         */
        public void SetBrcTop(BorderCode field_12_brcTop)
        {
            this.field_12_brcTop = field_12_brcTop;
        }

        /**
         * Get the brcLeft field for the TAP record.
         */
        public BorderCode GetBrcLeft()
        {
            return field_13_brcLeft;
        }

        /**
         * Set the brcLeft field for the TAP record.
         */
        public void SetBrcLeft(BorderCode field_13_brcLeft)
        {
            this.field_13_brcLeft = field_13_brcLeft;
        }

        /**
         * Get the brcRight field for the TAP record.
         */
        public BorderCode GetBrcRight()
        {
            return field_14_brcRight;
        }

        /**
         * Set the brcRight field for the TAP record.
         */
        public void SetBrcRight(BorderCode field_14_brcRight)
        {
            this.field_14_brcRight = field_14_brcRight;
        }

        /**
         * Get the brcVertical field for the TAP record.
         */
        public BorderCode GetBrcVertical()
        {
            return field_15_brcVertical;
        }

        /**
         * Set the brcVertical field for the TAP record.
         */
        public void SetBrcVertical(BorderCode field_15_brcVertical)
        {
            this.field_15_brcVertical = field_15_brcVertical;
        }

        /**
         * Get the brcHorizontal field for the TAP record.
         */
        public BorderCode GetBrcHorizontal()
        {
            return field_16_brcHorizontal;
        }

        /**
         * Set the brcHorizontal field for the TAP record.
         */
        public void SetBrcHorizontal(BorderCode field_16_brcHorizontal)
        {
            this.field_16_brcHorizontal = field_16_brcHorizontal;
        }


    }
}
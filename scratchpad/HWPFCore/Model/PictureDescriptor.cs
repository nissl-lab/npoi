/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;

namespace NPOI.HWPF.Model
{
    public class PictureDescriptor
    {
        private static int LCB_OFFSET = 0x00;
        private static int CBHEADER_OFFSET = 0x04;

        private static int MFP_MM_OFFSET = 0x06;
        private static int MFP_XEXT_OFFSET = 0x08;
        private static int MFP_YEXT_OFFSET = 0x0A;
        private static int MFP_HMF_OFFSET = 0x0C;

        private static int DXAGOAL_OFFSET = 0x1C;
        private static int DYAGOAL_OFFSET = 0x1E;

        private static int MX_OFFSET = 0x20;
        private static int MY_OFFSET = 0x22;

        private static int DXACROPLEFT_OFFSET = 0x24;
        private static int DYACROPTOP_OFFSET = 0x26;
        private static int DXACROPRIGHT_OFFSET = 0x28;
        private static int DYACROPBOTTOM_OFFSET = 0x2A;

        /**
         * Number of bytes in the PIC structure plus size of following picture data
         * which may be a Window's metafile, a bitmap, or the filename of a TIFF
         * file. In the case of a Macintosh PICT picture, this includes the size of
         * the PIC, the standard "x" metafile, and the Macintosh PICT data. See
         * Appendix B for more information.
         */
        protected int lcb;

        /**
         * Number of bytes in the PIC (to allow for future expansion).
         */
        protected int cbHeader;

        /*
         * Microsoft Office Word 97-2007 Binary File Format (.doc) Specification
         * 
         * Page 181 of 210
         * 
         * If a Windows metafile is stored immediately following the PIC structure,
         * the mfp is a Window's METAFILEPICT structure. See
         * http://msdn2.microsoft.com/en-us/library/ms649017(VS.85).aspx for more
         * information about the METAFILEPICT structure and
         * http://download.microsoft.com/download/0/B/E/0BE8BDD7-E5E8-422A-ABFD-
         * 4342ED7AD886/WindowsMetafileFormat(wmf)Specification.pdf for Windows
         * Metafile Format specification.
         * 
         * When the data immediately following the PIC is a TIFF filename,
         * mfp.mm==98 If a bitmap is stored after the pic, mfp.mm==99.
         * 
         * When the PIC describes a bitmap, mfp.xExt is the width of the bitmap in
         * pixels and mfp.yExt is the height of the bitmap in pixels.
         */

        protected int mfp_mm;
        protected int mfp_xExt;
        protected int mfp_yExt;
        protected int mfp_hMF;

        /**
         * <li>Window's bitmap structure when PIC describes a BITMAP (14 bytes)
         * 
         * <li>Rectangle for window origin and extents when metafile is stored --
         * ignored if 0 (8 bytes)
         */
        protected byte[] offset14 = new byte[14];

        /**
         * Horizontal measurement in twips of the rectangle the picture should be
         * imaged within
         */
        protected short dxaGoal = 0;

        /**
         * Vertical measurement in twips of the rectangle the picture should be
         * imaged within
         */
        protected short dyaGoal = 0;

        /**
         * Horizontal scaling factor supplied by user expressed in .001% units
         */
        protected short mx;

        /**
         * Vertical scaling factor supplied by user expressed in .001% units
         */
        protected short my;

        /**
         * The amount the picture has been cropped on the left in twips
         */
        protected short dxaCropLeft = 0;

        /**
         * The amount the picture has been cropped on the top in twips
         */
        protected short dyaCropTop = 0;

        /**
         * The amount the picture has been cropped on the right in twips
         */
        protected short dxaCropRight = 0;

        /**
         * The amount the picture has been cropped on the bottom in twips
         */
        protected short dyaCropBottom = 0;

        public PictureDescriptor()
        {
        }

        public PictureDescriptor(byte[] _dataStream, int startOffset)
        {
            this.lcb = LittleEndian.GetInt(_dataStream, startOffset + LCB_OFFSET);
            this.cbHeader = LittleEndian.GetUShort(_dataStream, startOffset
                    + CBHEADER_OFFSET);

            this.mfp_mm = LittleEndian.GetUShort(_dataStream, startOffset
                    + MFP_MM_OFFSET);
            this.mfp_xExt = LittleEndian.GetUShort(_dataStream, startOffset
                    + MFP_XEXT_OFFSET);
            this.mfp_yExt = LittleEndian.GetUShort(_dataStream, startOffset
                    + MFP_YEXT_OFFSET);
            this.mfp_hMF = LittleEndian.GetUShort(_dataStream, startOffset
                    + MFP_HMF_OFFSET);

            this.offset14 = LittleEndian.GetByteArray(_dataStream,
                    startOffset + 0x0E, 14);

            this.dxaGoal = LittleEndian.GetShort(_dataStream, startOffset
                    + DXAGOAL_OFFSET);
            this.dyaGoal = LittleEndian.GetShort(_dataStream, startOffset
                    + DYAGOAL_OFFSET);

            this.mx = LittleEndian.GetShort(_dataStream, startOffset + MX_OFFSET);
            this.my = LittleEndian.GetShort(_dataStream, startOffset + MY_OFFSET);

            this.dxaCropLeft = LittleEndian.GetShort(_dataStream, startOffset
                    + DXACROPLEFT_OFFSET);
            this.dyaCropTop = LittleEndian.GetShort(_dataStream, startOffset
                    + DYACROPTOP_OFFSET);
            this.dxaCropRight = LittleEndian.GetShort(_dataStream, startOffset
                    + DXACROPRIGHT_OFFSET);
            this.dyaCropBottom = LittleEndian.GetShort(_dataStream, startOffset
                    + DYACROPBOTTOM_OFFSET);
        }

        public override String ToString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("[PICF]\n");
            stringBuilder.Append("        lcb           = ").Append(this.lcb)
                    .Append('\n');
            stringBuilder.Append("        cbHeader      = ")
                    .Append(this.cbHeader).Append('\n');

            stringBuilder.Append("        mfp.mm        = ").Append(this.mfp_mm)
                    .Append('\n');
            stringBuilder.Append("        mfp.xExt      = ")
                    .Append(this.mfp_xExt).Append('\n');
            stringBuilder.Append("        mfp.yExt      = ")
                    .Append(this.mfp_yExt).Append('\n');
            stringBuilder.Append("        mfp.hMF       = ")
                    .Append(this.mfp_hMF).Append('\n');

            stringBuilder.Append("        offset14      = ");
            foreach (byte b in this.offset14)
                stringBuilder.Append(b);
            stringBuilder.Append('\n');
            stringBuilder.Append("        dxaGoal       = ")
                    .Append(this.dxaGoal).Append('\n');
            stringBuilder.Append("        dyaGoal       = ")
                    .Append(this.dyaGoal).Append('\n');

            stringBuilder.Append("        dxaCropLeft   = ")
                    .Append(this.dxaCropLeft).Append('\n');
            stringBuilder.Append("        dyaCropTop    = ")
                    .Append(this.dyaCropTop).Append('\n');
            stringBuilder.Append("        dxaCropRight  = ")
                    .Append(this.dxaCropRight).Append('\n');
            stringBuilder.Append("        dyaCropBottom = ")
                    .Append(this.dyaCropBottom).Append('\n');

            stringBuilder.Append("[/PICF]");
            return stringBuilder.ToString();
        }
    }
}

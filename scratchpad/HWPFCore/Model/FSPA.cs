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

namespace NPOI.HWPF.Model
{
    using NPOI.Util;
    using System.Text;
    using System;


    /**
     * File Shape Address structure
     *
     * @author Squeeself
     */
    public class FSPA
    {
        public static int FSPA_SIZE = 26;
        private int spid; // Shape identifier. Used to get data position
        private int xaLeft; // Enclosing rectangle
        private int yaTop; // Enclosing rectangle
        private int xaRight; // Enclosing rectangle
        private int yaBottom; // Enclosing rectangle
        private short options;
        private static BitField fHdr = BitFieldFactory.GetInstance(0x0001); // 1 in undo when in header
        private static BitField bx = BitFieldFactory.GetInstance(0x0006); // x pos relative to anchor CP: 0 - page margin, 1 - top of page, 2 - text, 3 - reserved
        private static BitField by = BitFieldFactory.GetInstance(0x0018); // y pos relative to anchor CP: ditto
        private static BitField wr = BitFieldFactory.GetInstance(0x01E0); // Text wrapping mode: 0 - like 2 w/o absolute, 1 - no text next to shape, 2 - wrap around absolute object, 3 - wrap as if no object, 4 - wrap tightly around object, 5 - wrap tightly, allow holes, 6-15 - reserved
        private static BitField wrk = BitFieldFactory.GetInstance(0x1E00); // Text wrapping mode type (for modes 2&4): 0 - wrap both sides, 1 - wrap only left, 2 - wrap only right, 3 - wrap largest side
        private static BitField fRcaSimple = BitFieldFactory.GetInstance(0x2000); // OverWrites bx if Set, forcing rectangle to be page relative
        private static BitField fBelowText = BitFieldFactory.GetInstance(0x4000); // if true, shape is below text, otherwise above
        private static BitField fAnchorLock = BitFieldFactory.GetInstance(0x8000); // if true, anchor is locked
        private int cTxbx; // Count of textboxes in shape (undo doc only)

        public FSPA()
        {
        }

        public FSPA(byte[] bytes, int offset)
        {
            spid = LittleEndian.GetInt(bytes, offset);
            offset += LittleEndianConsts.INT_SIZE;
            xaLeft = LittleEndian.GetInt(bytes, offset);
            offset += LittleEndianConsts.INT_SIZE;
            yaTop = LittleEndian.GetInt(bytes, offset);
            offset += LittleEndianConsts.INT_SIZE;
            xaRight = LittleEndian.GetInt(bytes, offset);
            offset += LittleEndianConsts.INT_SIZE;
            yaBottom = LittleEndian.GetInt(bytes, offset);
            offset += LittleEndianConsts.INT_SIZE;
            options = LittleEndian.GetShort(bytes, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            cTxbx = LittleEndian.GetInt(bytes, offset);
        }

        public int GetSpid()
        {
            return spid;
        }

        public int GetXaLeft()
        {
            return xaLeft;
        }

        public int GetYaTop()
        {
            return yaTop;
        }

        public int GetXaRight()
        {
            return xaRight;
        }

        public int GetYaBottom()
        {
            return yaBottom;
        }

        public bool IsFHdr()
        {
            return fHdr.IsSet(options);
        }

        public short GetBx()
        {
            return bx.GetShortValue(options);
        }

        public short GetBy()
        {
            return by.GetShortValue(options);
        }

        public short GetWr()
        {
            return wr.GetShortValue(options);
        }

        public short GetWrk()
        {
            return wrk.GetShortValue(options);
        }

        public bool IsFRcaSimple()
        {
            return fRcaSimple.IsSet(options);
        }

        public bool IsFBelowText()
        {
            return fBelowText.IsSet(options);
        }

        public bool IsFAnchorLock()
        {
            return fAnchorLock.IsSet(options);
        }

        public int GetCTxbx()
        {
            return cTxbx;
        }

        public byte[] ToArray()
        {
            int offset = 0;
            byte[] buf = new byte[FSPA_SIZE];

            LittleEndian.PutInt(buf, offset, spid);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, xaLeft);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, yaTop);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, xaRight);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, yaBottom);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutShort(buf, offset, options);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutInt(buf, offset, cTxbx);
            offset += LittleEndianConsts.INT_SIZE;

            return buf;
        }

        public override String ToString()
        {
            StringBuilder buf = new StringBuilder();
            buf.Append("spid: ").Append(spid);
            buf.Append(", xaLeft: ").Append(xaLeft);
            buf.Append(", yaTop: ").Append(yaTop);
            buf.Append(", xaRight: ").Append(xaRight);
            buf.Append(", yaBottom: ").Append(yaBottom);
            buf.Append(", options: ").Append(options);
            buf.Append(" (fHdr: ").Append(IsFHdr());
            buf.Append(", bx: ").Append(GetBx());
            buf.Append(", by: ").Append(GetBy());
            buf.Append(", wr: ").Append(GetWr());
            buf.Append(", wrk: ").Append(GetWrk());
            buf.Append(", fRcaSimple: ").Append(IsFRcaSimple());
            buf.Append(", fBelowText: ").Append(IsFBelowText());
            buf.Append(", fAnchorLock: ").Append(IsFAnchorLock());
            buf.Append("), cTxbx: ").Append(cTxbx);
            return buf.ToString();
        }
    }



}
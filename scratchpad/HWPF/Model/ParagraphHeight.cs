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
    using System;
    using NPOI.Util;
    using System.IO;

    public class ParagraphHeight
    {
        private short infoField;
        private BitField fSpare = BitFieldFactory.GetInstance(0x0001);
        private BitField fUnk = BitFieldFactory.GetInstance(0x0002);
        private BitField fDiffLines = BitFieldFactory.GetInstance(0x0004);
        private BitField clMac = BitFieldFactory.GetInstance(0xff00);
        private short reserved;
        private int dxaCol;
        private int dymLineOrHeight;

        public ParagraphHeight(byte[] buf, int offset)
        {
            infoField = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            reserved = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            dxaCol = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            dymLineOrHeight = LittleEndian.GetInt(buf, offset);
        }

        public ParagraphHeight()
        {

        }

        public void Write(Stream out1)
        {
            byte[] bytes=ToArray();
            out1.Write(bytes, (int)out1.Position, bytes.Length);
        }

        internal byte[] ToArray()
        {
            byte[] buf = new byte[12];
            int offset = 0;
            LittleEndian.PutShort(buf, offset, infoField);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, reserved);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutInt(buf, offset, dxaCol);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, dymLineOrHeight);

            return buf;
        }

        public override bool Equals(Object o)
        {
            ParagraphHeight ph = (ParagraphHeight)o;

            return infoField == ph.infoField && reserved == ph.reserved &&
                   dxaCol == ph.dxaCol && dymLineOrHeight == ph.dymLineOrHeight;
        }
    }
}


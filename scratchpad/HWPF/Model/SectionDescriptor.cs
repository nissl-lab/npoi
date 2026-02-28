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

    public class SectionDescriptor
    {
        /// <summary>
        /// Used internally by Word
        /// </summary>
        private short fn;
        /// <summary>
        /// File offset in main stream to beginning of SEPX stored for section. 
        /// If sed.fcSepx==0xFFFFFFFF, the section properties for the section are equal
        /// to the standard SEP (see SEP definition).
        /// </summary>
        private int fcSepx;
        /// <summary>
        /// Used internally by Word
        /// </summary>
        private short fnMpr;
        /// <summary>
        /// Points to offset in FC space of main stream where the Macintosh Print
        /// Record for a document created on a Macintosh will be stored
        /// </summary>
        private int fcMpr;

        public SectionDescriptor()
        {
        }

        public SectionDescriptor(byte[] buf, int offset)
        {
            fn = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            fcSepx = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            fnMpr = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            fcMpr = LittleEndian.GetInt(buf, offset);
        }

        public int GetFc()
        {
            return fcSepx;
        }

        public void SetFc(int fc)
        {
            this.fcSepx = fc;
        }

        public override bool Equals(Object o)
        {
            SectionDescriptor sed = (SectionDescriptor)o;
            return sed.fn == fn && sed.fnMpr == fnMpr;
        }

        public byte[] ToArray()
        {
            int offset = 0;
            byte[] buf = new byte[12];

            LittleEndian.PutShort(buf, offset, fn);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutInt(buf, offset, fcSepx);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutShort(buf, offset, fnMpr);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutInt(buf, offset, fcMpr);

            return buf;
        }
        public override string ToString()
        {

            return "[SED] (fn: " + fn + "; fcSepx: " + fcSepx + "; fnMpr: " + fnMpr
                + "; fcMpr: " + fcMpr + ")";
        }
    }
}


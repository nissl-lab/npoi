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
    using NPOI.HWPF.SPRM;
    using NPOI.Util;
    using System;
    using NPOI.HWPF.UserModel;


    /**
     * DANGER - works in bytes!
     *
     * Make sure you call GetStart() / GetEnd() when you want characters
     *  (normal use), but GetStartByte() / GetEndByte() when you're
     *  Reading in / writing out!
     *
     * @author Ryan Ackley
     */

    public class PAPX : BytePropertyNode
    {

        private ParagraphHeight _phe;

        public PAPX(int charStart, int charEnd, SprmBuffer buf):
            base(charStart, charEnd, buf)
        {
            
            _phe = new ParagraphHeight();
        }

        public PAPX(int fcStart, int fcEnd, CharIndexTranslator translator, byte[] papx, ParagraphHeight phe, byte[] dataStream)
            : base(fcStart, fcEnd, translator, new SprmBuffer(papx,2))
        {

            _phe = phe;
            SprmBuffer buf = FindHuge(new SprmBuffer(papx,2), dataStream);
            if (buf != null)
                _buf = buf;
        }
        public PAPX(int charStart, int charEnd, byte[] papx, ParagraphHeight phe,
                    byte[] dataStream):
            base(charStart, charEnd, new SprmBuffer(papx, 2))
        {
            _phe = phe;
            SprmBuffer buf = FindHuge(new SprmBuffer(papx, 2), dataStream);
            if (buf != null)
                _buf = buf;
        }
        [Obsolete]
        public PAPX(int fcStart, int fcEnd, CharIndexTranslator translator, SprmBuffer buf, byte[] dataStream)
            : base(fcStart, fcEnd, translator, buf)
        {

            _phe = new ParagraphHeight();
            buf = FindHuge(buf, dataStream);
            if (buf != null)
                _buf = buf;
        }

        private SprmBuffer FindHuge(SprmBuffer buf, byte[] datastream)
        {
            byte[] grpprl = buf.ToByteArray();
            if (grpprl.Length == 8 && datastream != null) // then check for sprmPHugePapx
            {
                SprmOperation sprm = new SprmOperation(grpprl, 2);
                if ((sprm.Operation == 0x45 || sprm.Operation == 0x46)
                    && sprm.SizeCode == 3)
                {
                    int hugeGrpprlOffset = sprm.Operand;
                    if (hugeGrpprlOffset + 1 < datastream.Length)
                    {
                        int grpprlSize = LittleEndian.GetShort(datastream, hugeGrpprlOffset);
                        if (hugeGrpprlOffset + grpprlSize < datastream.Length)
                        {
                            byte[] hugeGrpprl = new byte[grpprlSize + 2];
                            // copy original istd into huge Grpprl
                            hugeGrpprl[0] = grpprl[0]; 
                            hugeGrpprl[1] = grpprl[1];
                            // copy Grpprl from dataStream
                            Array.Copy(datastream, hugeGrpprlOffset + 2, hugeGrpprl, 2,
                                             grpprlSize);

                            return new SprmBuffer(hugeGrpprl,2);
                        }
                    }
                }
            }
            return null;
        }


        public ParagraphHeight GetParagraphHeight()
        {
            return _phe;
        }

        public byte[] GetGrpprl()
        {
            return ((SprmBuffer)_buf).ToByteArray();
        }

        public short GetIstd()
        {
            byte[] buf = GetGrpprl();
            if (buf.Length == 0)
            {
                return 0;
            }
            if (buf.Length == 1)
            {
                return (short)LittleEndian.GetUByte(buf, 0);
            }
            return LittleEndian.GetShort(buf);
        }

        public SprmBuffer GetSprmBuf()
        {
            return (SprmBuffer)_buf;
        }

        public ParagraphProperties GetParagraphProperties(StyleSheet ss)
        {
            if (ss == null)
            {
                // TODO Fix up for Word 6/95
                return new ParagraphProperties();
            }

            short istd = GetIstd();
            ParagraphProperties baseStyle = ss.GetParagraphStyle(istd);
            ParagraphProperties props = ParagraphSprmUncompressor.UncompressPAP(baseStyle, GetGrpprl(), 2);
            return props;
        }

        public override bool Equals(Object o)
        {
            if (base.Equals(o))
            {
                return _phe.Equals(((PAPX)o)._phe);
            }
            return false;
        }
    }
}


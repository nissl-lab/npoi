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
    using System;
    using System.Text;
    using NPOI.HWPF.UserModel;
    /**
     * Comment me
     *
     * @author Ryan Ackley
     */

    public class StyleDescription
    {

        private const int PARAGRAPH_STYLE = 1;
        private const int CHARACTER_STYLE = 2;

        private int _istd;
        private int _baseLength;
        private short _infoshort;
        private static BitField _sti = BitFieldFactory.GetInstance(0xfff);
        private static BitField _fScratch = BitFieldFactory.GetInstance(0x1000);
        private static BitField _fInvalHeight = BitFieldFactory.GetInstance(0x2000);
        private static BitField _fHasUpe = BitFieldFactory.GetInstance(0x4000);
        private static BitField _fMassCopy = BitFieldFactory.GetInstance(0x8000);
        private short _infoshort2;
        private static BitField _styleTypeCode = BitFieldFactory.GetInstance(0xf);
        private static BitField _baseStyle = BitFieldFactory.GetInstance(0xfff0);
        private short _infoshort3;
        private static BitField _numUPX = BitFieldFactory.GetInstance(0xf);
        private static BitField _nextStyle = BitFieldFactory.GetInstance(0xfff0);
        private short _bchUpe;
        private short _infoshort4;
        private static BitField _fAutoRedef = BitFieldFactory.GetInstance(0x1);
        private static BitField _fHidden = BitFieldFactory.GetInstance(0x2);

        UPX[] _upxs;
        String _name;
        ParagraphProperties _pap;
        CharacterProperties _chp;

        public StyleDescription()
        {
            //      _pap = new ParagraphProperties();
            //      _chp = new CharacterProperties();
        }
        public StyleDescription(byte[] std, int baseLength, int offset, bool word9)
        {
            _baseLength = baseLength;
            int nameStart = offset + baseLength;
            _infoshort = LittleEndian.GetShort(std, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _infoshort2 = LittleEndian.GetShort(std, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _infoshort3 = LittleEndian.GetShort(std, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _bchUpe = LittleEndian.GetShort(std, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _infoshort4 = LittleEndian.GetShort(std, offset);
            offset += LittleEndianConsts.SHORT_SIZE;

            //first byte(s) of variable length section of std is the length of the
            //style name and aliases string
            int nameLength = 0;
            int multiplier = 1;
            if (word9)
            {
                nameLength = LittleEndian.GetShort(std, nameStart);
                multiplier = 2;
                nameStart += LittleEndianConsts.SHORT_SIZE;
            }
            else
            {
                nameLength = std[nameStart];
            }

            try
            {
                _name = Encoding.GetEncoding("UTF-16LE").GetString(std, nameStart, nameLength * multiplier);
            }
            catch (EncoderFallbackException)
            {
                // ignore
            }

            //length then null terminator.
            int grupxStart = ((nameLength + 1) * multiplier) + nameStart;

            // the spec only refers to two possible upxs but it mentions
            // that more may be Added in the future
            int varoffset = grupxStart;
            int numUPX = _numUPX.GetValue(_infoshort3);
            _upxs = new UPX[numUPX];
            for (int x = 0; x < numUPX; x++)
            {
                int upxSize = LittleEndian.GetShort(std, varoffset);
                varoffset += LittleEndianConsts.SHORT_SIZE;

                byte[] upx = new byte[upxSize];
                Array.Copy(std, varoffset, upx, 0, upxSize);
                _upxs[x] = new UPX(upx);
                varoffset += upxSize;


                // the upx will always start on a word boundary.
                if (upxSize % 2 == 1)
                {
                    ++varoffset;
                }

            }


        }
        public int GetBaseStyle()
        {
            return _baseStyle.GetValue(_infoshort2);
        }
        public byte[] GetCHPX()
        {
            switch (_styleTypeCode.GetValue(_infoshort2))
            {
                case PARAGRAPH_STYLE:
                    if (_upxs.Length > 1)
                    {
                        return _upxs[1].GetUPX();
                    }
                    return null;
                case CHARACTER_STYLE:
                    return _upxs[0].GetUPX();
                default:
                    return null;
            }

        }
        public byte[] GetPAPX()
        {
            switch (_styleTypeCode.GetValue(_infoshort2))
            {
                case PARAGRAPH_STYLE:
                    return _upxs[0].GetUPX();
                default:
                    return null;
            }
        }
        public ParagraphProperties GetPAP()
        {
            return _pap;
        }
        public CharacterProperties GetCHP()
        {
            return _chp;
        }
        internal void SetPAP(ParagraphProperties pap)
        {
            _pap = pap;
        }
        internal  void SetCHP(CharacterProperties chp)
        {
            _chp = chp;
        }

        public String GetName()
        {
            return _name;
        }

        public byte[] ToArray()
        {
            // size Equals _baseLength bytes for known variables plus 2 bytes for name
            // length plus name length * 2 plus 2 bytes for null plus upx's preceded by
            // length
            int size = _baseLength + 2 + ((_name.Length + 1) * 2);

            // determine the size needed for the upxs. They always fall on word
            // boundaries.
            size += _upxs[0].Size + 2;
            for (int x = 1; x < _upxs.Length; x++)
            {
                size += _upxs[x - 1].Size % 2;
                size += _upxs[x].Size + 2;
            }


            byte[] buf = new byte[size];

            int offset = 0;
            LittleEndian.PutShort(buf, offset, _infoshort);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, _infoshort2);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, _infoshort3);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, _bchUpe);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, _infoshort4);
            offset = _baseLength;

            char[] letters = _name.ToCharArray();
            LittleEndian.PutShort(buf, _baseLength, (short)letters.Length);
            offset += LittleEndianConsts.SHORT_SIZE;
            for (int x = 0; x < letters.Length; x++)
            {
                LittleEndian.PutShort(buf, offset, (short)letters[x]);
                offset += LittleEndianConsts.SHORT_SIZE;
            }
            // get past the null delimiter for the name.
            offset += LittleEndianConsts.SHORT_SIZE;

            for (int x = 0; x < _upxs.Length; x++)
            {
                short upxSize = (short)_upxs[x].Size;
                LittleEndian.PutShort(buf, offset, upxSize);
                offset += LittleEndianConsts.SHORT_SIZE;
                Array.Copy(_upxs[x].GetUPX(), 0, buf, offset, upxSize);
                offset += upxSize + (upxSize % 2);
            }

            return buf;
        }

        public override bool Equals(Object o)
        {
            StyleDescription sd = (StyleDescription)o;
            if (sd._infoshort == _infoshort && sd._infoshort2 == _infoshort2 &&
                sd._infoshort3 == _infoshort3 && sd._bchUpe == _bchUpe &&
                sd._infoshort4 == _infoshort4 &&
                _name.Equals(sd._name))
            {

                if (!Arrays.Equals(_upxs, sd._upxs))
                {
                    return false;
                }
                return true;
            }
            return false;
        }
    }

}

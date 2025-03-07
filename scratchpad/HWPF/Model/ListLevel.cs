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

    /**
     *
     */
    public class ListLevel
    {
        private static int RGBXCH_NUMS_SIZE = 9;

        private int _iStartAt;
        private byte _nfc;
        private byte _info;
        private static BitField _jc;
        private static BitField _fLegal;
        private static BitField _fNoRestart;
        private static BitField _fPrev;
        private static BitField _fPrevSpace;
        private static BitField _fWord6;
        private byte[] _rgbxchNums;
        private byte _ixchFollow;
        private int _dxaSpace;
        private int _dxaIndent;
        private int _cbGrpprlChpx;
        private int _cbGrpprlPapx;
        private short _reserved;
        private byte[] _grpprlPapx;
        private byte[] _grpprlChpx;
        private char[] _numberText = null;

        public ListLevel(int startAt, int numberFormatCode, int alignment,
                         byte[] numberProperties, byte[] entryProperties,
                         String numberText)
        {
            _iStartAt = startAt;
            _nfc = (byte)numberFormatCode;
            _jc.SetValue(_info, alignment);
            _grpprlChpx = numberProperties;
            _grpprlPapx = entryProperties;
            _numberText = numberText.ToCharArray();
        }

        public ListLevel(int level, bool numbered)
        {
            _iStartAt = 1;
            _grpprlPapx = Array.Empty<byte>();
            _grpprlChpx = Array.Empty<byte>();
            _numberText = Array.Empty<char>();
            _rgbxchNums = new byte[RGBXCH_NUMS_SIZE];

            if (numbered)
            {
                _rgbxchNums[0] = 1;
                _numberText = new char[] { (char)level, '.' };
            }
            else
            {
                _numberText = new char[] { '\u2022' };
            }
        }

        public ListLevel(byte[] buf, int offset)
        {
            _iStartAt = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _nfc = buf[offset++];
            _info = buf[offset++];

            _rgbxchNums = new byte[RGBXCH_NUMS_SIZE];
            Array.Copy(buf, offset, _rgbxchNums, 0, RGBXCH_NUMS_SIZE);
            offset += RGBXCH_NUMS_SIZE;

            _ixchFollow = buf[offset++];
            _dxaSpace = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _dxaIndent = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _cbGrpprlChpx = LittleEndian.GetUByte(buf, offset++);
            _cbGrpprlPapx = LittleEndian.GetUByte(buf, offset++);
            _reserved = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;

            _grpprlPapx = new byte[_cbGrpprlPapx];
            _grpprlChpx = new byte[_cbGrpprlChpx];
            Array.Copy(buf, offset, _grpprlPapx, 0, _cbGrpprlPapx);
            offset += _cbGrpprlPapx;
            Array.Copy(buf, offset, _grpprlChpx, 0, _cbGrpprlChpx);
            offset += _cbGrpprlChpx;

            int numberTextLength = LittleEndian.GetShort(buf, offset);
            /* sometimes numberTextLength<0 */
            /* by derjohng */
            if (numberTextLength > 0)
            {
                _numberText = new char[numberTextLength];
                offset += LittleEndianConsts.SHORT_SIZE;
                for (int x = 0; x < numberTextLength; x++)
                {
                    _numberText[x] = (char)LittleEndian.GetShort(buf, offset);
                    offset += LittleEndianConsts.SHORT_SIZE;
                }
            }

        }

        public int GetStartAt()
        {
            return _iStartAt;
        }

        public int GetNumberFormat()
        {
            return _nfc;
        }

        public int GetAlignment()
        {
            return _jc.GetValue(_info);
        }

        public String GetNumberText()
        {
            if (_numberText != null)
                return new String(_numberText);
            else
                return null;
        }
        /**
     * "The type of character following the number text for the paragraph: 0 == tab, 1 == space, 2 == nothing."
     */
        public byte GetTypeOfCharFollowingTheNumber()
        {
            return this._ixchFollow;
        }
        public void SetStartAt(int startAt)
        {
            _iStartAt = startAt;
        }

        public void SetNumberFormat(int numberFormatCode)
        {
            _nfc = (byte)numberFormatCode;
        }

        public void SetAlignment(int alignment)
        {
            _jc.SetValue(_info, alignment);
        }

        public void SetNumberProperties(byte[] grpprl)
        {
            _grpprlChpx = grpprl;

        }

        public void SetLevelProperties(byte[] grpprl)
        {
            _grpprlPapx = grpprl;
        }

        public byte[] GetLevelProperties()
        {
            return _grpprlPapx;
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
            {
                return false;
            }

            ListLevel lvl = (ListLevel)obj;
            return _cbGrpprlChpx == lvl._cbGrpprlChpx && lvl._cbGrpprlPapx == _cbGrpprlPapx &&
              lvl._dxaIndent == _dxaIndent && lvl._dxaSpace == _dxaSpace &&
              Arrays.Equals(lvl._grpprlChpx, _grpprlChpx) &&
              Arrays.Equals(lvl._grpprlPapx, _grpprlPapx) &&
              lvl._info == _info && lvl._iStartAt == _iStartAt &&
              lvl._ixchFollow == _ixchFollow && lvl._nfc == _nfc &&
              Arrays.Equals(lvl._numberText, _numberText) &&
              Arrays.Equals(lvl._rgbxchNums, _rgbxchNums) &&
              lvl._reserved == _reserved;


        }
        public byte[] ToArray()
        {
            byte[] buf = new byte[GetSizeInBytes()];
            int offset = 0;
            LittleEndian.PutInt(buf, offset, _iStartAt);
            offset += LittleEndianConsts.INT_SIZE;
            buf[offset++] = _nfc;
            buf[offset++] = _info;
            Array.Copy(_rgbxchNums, 0, buf, offset, RGBXCH_NUMS_SIZE);
            offset += RGBXCH_NUMS_SIZE;
            buf[offset++] = _ixchFollow;
            LittleEndian.PutInt(buf, offset, _dxaSpace);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, _dxaIndent);
            offset += LittleEndianConsts.INT_SIZE;

            buf[offset++] = (byte)_cbGrpprlChpx;
            buf[offset++] = (byte)_cbGrpprlPapx;
            LittleEndian.PutShort(buf, offset, _reserved);
            offset += LittleEndianConsts.SHORT_SIZE;

            Array.Copy(_grpprlPapx, 0, buf, offset, _cbGrpprlPapx);
            offset += _cbGrpprlPapx;
            Array.Copy(_grpprlChpx, 0, buf, offset, _cbGrpprlChpx);
            offset += _cbGrpprlChpx;

            if (_numberText == null)
            {
                // TODO - write junit to test this flow
                LittleEndian.PutUShort(buf, offset, 0);
            }
            else
            {
                LittleEndian.PutUShort(buf, offset, _numberText.Length);
                offset += LittleEndianConsts.SHORT_SIZE;
                for (int x = 0; x < _numberText.Length; x++)
                {
                    LittleEndian.PutUShort(buf, offset, _numberText[x]);
                    offset += LittleEndianConsts.SHORT_SIZE;
                }
            }
            return buf;
        }
        public int GetSizeInBytes()
        {
            int result =
                6 // int byte byte
                + RGBXCH_NUMS_SIZE
                + 13 // byte int int byte byte short
                + _cbGrpprlChpx
                + _cbGrpprlPapx
                + 2; // numberText length
            if (_numberText != null)
            {
                result += _numberText.Length * LittleEndianConsts.SHORT_SIZE;
            }
            return result;
        }

    }
}


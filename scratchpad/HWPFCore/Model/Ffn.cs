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
     * FFN - Font Family Name. FFN is a data structure that stores the names of the Main
     * Font and that of Alternate font as an array of characters. It has also a header
     * that stores info about the whole structure and the fonts
     *
     * @author Praveen Mathew
     */
    public class Ffn
    {
        private int _cbFfnM1;//total length of FFN - 1.
        private byte _info;
        private static BitField _prq = BitFieldFactory.GetInstance(0x0003);// pitch request
        private static BitField _fTrueType = BitFieldFactory.GetInstance(0x0004);// when 1, font is a TrueType font
        private static BitField _ff = BitFieldFactory.GetInstance(0x0070);
        private short _wWeight;// base weight of font
        private byte _chs;// character set identifier
        private byte _ixchSzAlt;  // index into ffn.szFfn to the name of
        // the alternate font
        private byte[] _panose = new byte[10];//????
        private byte[] _fontSig = new byte[24];//????

        // zero terminated string that records name of font, cuurently not
        // supporting Extended chars
        private char[] _xszFfn;

        // extra facilitator members
        private int _xszFfnLength;

        public Ffn(byte[] buf, int offset)
        {
            int offsetTmp = offset;

            _cbFfnM1 = LittleEndian.GetUByte(buf, offset);
            offset += LittleEndianConsts.BYTE_SIZE;
            _info = buf[offset];
            offset += LittleEndianConsts.BYTE_SIZE;
            _wWeight = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _chs = buf[offset];
            offset += LittleEndianConsts.BYTE_SIZE;
            _ixchSzAlt = buf[offset];
            offset += LittleEndianConsts.BYTE_SIZE;

            // read panose and fs so we can write them back out.
            Array.Copy(buf, offset, _panose, 0, _panose.Length);
            offset += _panose.Length;
            Array.Copy(buf, offset, _fontSig, 0, _fontSig.Length);
            offset += _fontSig.Length;

            offsetTmp = offset - offsetTmp;
            _xszFfnLength = (this.GetSize() - offsetTmp) / 2;
            _xszFfn = new char[_xszFfnLength];

            for (int i = 0; i < _xszFfnLength; i++)
            {
                _xszFfn[i] = (char)LittleEndian.GetShort(buf, offset);
                offset += LittleEndianConsts.SHORT_SIZE;
            }


        }

        public int Get_cbFfnM1()
        {
            return _cbFfnM1;
        }

        public short GetWeight()
        {
            return _wWeight;
        }

        public byte GetChs()
        {
            return _chs;
        }

        public byte[] GetPanose()
        {
            return _panose;
        }

        public byte[] GetFontSig()
        {
            return _fontSig;
        }

        public int GetSize()
        {
            return (_cbFfnM1 + 1);
        }

        public String GetMainFontName()
        {
            int index = 0;
            for (; index < _xszFfnLength; index++)
            {
                if (_xszFfn[index] == '\0')
                {
                    break;
                }
            }
            return new String(_xszFfn, 0, index);
        }

        public String GetAltFontName()
        {
            int index = _ixchSzAlt;
            for (; index < _xszFfnLength; index++)
            {
                if (_xszFfn[index] == '\0')
                {
                    break;
                }
            }
            return new String(_xszFfn, _ixchSzAlt, index);

        }

        public void Set_cbFfnM1(int _cbFfnM1)
        {
            this._cbFfnM1 = _cbFfnM1;
        }

        // Changed protected to public
        public byte[] ToArray()
        {
            int offset = 0;
            byte[] buf = new byte[this.GetSize()];

            buf[offset] = (byte)_cbFfnM1;
            offset += LittleEndianConsts.BYTE_SIZE;
            buf[offset] = _info;
            offset += LittleEndianConsts.BYTE_SIZE;
            LittleEndian.PutShort(buf, offset, _wWeight);
            offset += LittleEndianConsts.SHORT_SIZE;
            buf[offset] = _chs;
            offset += LittleEndianConsts.BYTE_SIZE;
            buf[offset] = _ixchSzAlt;
            offset += LittleEndianConsts.BYTE_SIZE;

            Array.Copy(_panose, 0, buf, offset, _panose.Length);
            offset += _panose.Length;
            Array.Copy(_fontSig, 0, buf, offset, _fontSig.Length);
            offset += _fontSig.Length;

            for (int i = 0; i < _xszFfn.Length; i++)
            {
                LittleEndian.PutShort(buf, offset, (short)_xszFfn[i]);
                offset += LittleEndianConsts.SHORT_SIZE;
            }

            return buf;

        }

        public override bool Equals(Object o)
        {
            bool retVal = true;

            if (((Ffn)o).Get_cbFfnM1() == _cbFfnM1)
            {
                if (((Ffn)o)._info == _info)
                {
                    if (((Ffn)o)._wWeight == _wWeight)
                    {
                        if (((Ffn)o)._chs == _chs)
                        {
                            if (((Ffn)o)._ixchSzAlt == _ixchSzAlt)
                            {
                                if (Arrays.Equals(((Ffn)o)._panose, _panose))
                                {
                                    if (Arrays.Equals(((Ffn)o)._fontSig, _fontSig))
                                    {
                                        if (!(Arrays.Equals(((Ffn)o)._xszFfn, _xszFfn)))
                                            retVal = false;
                                    }
                                    else
                                        retVal = false;
                                }
                                else
                                    retVal = false;
                            }
                            else
                                retVal = false;
                        }
                        else
                            retVal = false;
                    }
                    else
                        retVal = false;
                }
                else
                    retVal = false;
            }
            else
                retVal = false;

            return retVal;
        }


    }

}



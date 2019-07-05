
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


namespace NPOI.HWPF.UserModel
{
    using System;
    using NPOI.Util;
    using NPOI.HWPF.Model;

    public class BorderCode:BaseObject
    {
        public static int SIZE = 4;
        private short _info;
        private static BitField _dptLineWidth = BitFieldFactory.GetInstance(0xff);
        private static BitField _brcType = BitFieldFactory.GetInstance(0xff00);
        private short _info2;
        private static BitField _ico = BitFieldFactory.GetInstance(0xff);
        private static BitField _dptSpace = BitFieldFactory.GetInstance(0x1f00);
        private static BitField _fShadow = BitFieldFactory.GetInstance(0x2000);
        private static BitField _fFrame = BitFieldFactory.GetInstance(0x4000);

        public BorderCode()
        {
        }

        public BorderCode(byte[] buf, int offset)
        {
            _info = LittleEndian.GetShort(buf, offset);
            _info2 = LittleEndian.GetShort(buf, offset + LittleEndianConsts.SHORT_SIZE);
        }

        public void Serialize(byte[] buf, int offset)
        {
            LittleEndian.PutShort(buf, offset, _info);
            LittleEndian.PutShort(buf, offset + LittleEndianConsts.SHORT_SIZE, _info2);
        }

        public int ToInt()
        {
            byte[] buf = new byte[4];
            Serialize(buf, 0);
            return LittleEndian.GetInt(buf);
        }

        public bool IsEmpty
        {
            get
            {
                return _info == 0 && _info2 == 0;
            }
        }

        public override bool Equals(Object o)
        {
            BorderCode brc = (BorderCode)o;
            return _info == brc._info && _info2 == brc._info2;
        }
        /**
         * Width of a single line in 1/8 pt, max of 32 pt.
         */
        public int LineWidth
        {
            get
            {
                return _dptLineWidth.GetShortValue(_info);
            }
            set 
            {
                _dptLineWidth.SetValue(_info, value);
            }
        }

        /**
         * Border type code:
         * <li>0  none
         * <li>1  single
         * <li>2  thick
         * <li>3  double
         * <li>5  hairline
         * <li>6  dot
         * <li>7  dash large gap
         * <li>8  dot dash
         * <li>9  dot dot dash
         * <li>10  triple
         * <li>11  thin-thick small gap
         * <li>12  thick-thin small gap
         * <li>13  thin-thick-thin small gap
         * <li>14  thin-thick medium gap
         * <li>15  thick-thin medium gap
         * <li>16  thin-thick-thin medium gap
         * <li>17  thin-thick large gap
         * <li>18  thick-thin large gap
         * <li>19  thin-thick-thin large gap
         * <li>20  wave
         * <li>21  double wave
         * <li>22  dash small gap
         * <li>23  dash dot stroked
         * <li>24  emboss 3D
         * <li>25  engrave 3D
         * <li>codes 64 - 230 represent border art types and are used only for page borders
         */
        public int BorderType
        {
            get
            {
                return _brcType.GetShortValue(_info);
            }
            set 
            {
                _brcType.SetValue(_info, value);
            }
        }

        /**
         * Color:
         * <li>0  Auto
         * <li>1  Black
         * <li>2  Blue
         * <li>3  Cyan
         * <li>4  Green
         * <li>5  Magenta
         * <li>6  Red
         * <li>7  Yellow
         * <li>8  White
         * <li>9  DkBlue
         * <li>10  DkCyan
         * <li>11  DkGreen
         * <li>12  DkMagenta
         * <li>13  DkRed
         * <li>14  DkYellow
         * <li>15  DkGray
         * <li>16  LtGray
         */
        public short Color
        {
            get
            {
                return _ico.GetShortValue(_info2);
            }
            set 
            {
                _ico.SetValue(_info2, value);
            }
        }

        /**
         * Width of space to maintain between border and text within border.
         * 
         * <p>Must be 0 when BRC is a substructure of TC.
         * 
         * <p>Stored in points.
         */
        public int Space
        {
            get
            {
                return _dptSpace.GetShortValue(_info2);
            }
            set 
            {
                _dptSpace.SetValue(_info2, value);
            }
        }

        /**
         * When true, border is drawn with shadow
         * Must be false when BRC is a substructure of the TC.
         */
        public bool IsShadow
        {
            get{
                return _fShadow.GetValue(_info2) != 0;
            }
            set 
            {
                _fShadow.SetValue(_info2, value ? 1 : 0);
            }
        }

        /**
         * Don't reverse the border.
         */
        public bool IsFrame
        {
            get
            {
                return _fFrame.GetValue(_info2) != 0;
            }
            set 
            {
                _fFrame.SetValue(_info2, value ? 1 : 0);
            }
        }

    }
}
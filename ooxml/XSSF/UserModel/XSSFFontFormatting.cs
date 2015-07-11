/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
namespace NPOI.XSSF.UserModel
{


    /**
     * @author Yegor Kozlov
     */
    public class XSSFFontFormatting : IFontFormatting
    {
        CT_Font _font;

        /*package*/
        internal XSSFFontFormatting(CT_Font font)
        {
            _font = font;
        }

        /**
         * Get the type of super or subscript for the font
         *
         * @return super or subscript option
         * @see #SS_NONE
         * @see #SS_SUPER
         * @see #SS_SUB
         */
        public FontSuperScript EscapementType
        {
            get
            {
                if (_font.sizeOfVertAlignArray() == 0) return FontSuperScript.None;

                CT_VerticalAlignFontProperty prop = _font.GetVertAlignArray(0);
                return (FontSuperScript)(prop.val - 1);
            }
            set 
            {
                _font.SetVertAlignArray(null);
                if (value != FontSuperScript.None)
                {
                    _font.AddNewVertAlign().val = (ST_VerticalAlignRun)(value + 1);
                }
            }
        }


        /**
         * @return font color index
         */
        public short FontColorIndex
        {
            get
            {
                if (_font.sizeOfColorArray() == 0) return -1;

                int idx = 0;
                CT_Color color = _font.GetColorArray(0);
                if (color.IsSetIndexed()) idx = (int)color.indexed;
                return (short)idx;
            }
            set 
            {
                _font.SetColorArray(null);
                if (value != -1)
                {
                    var clr=_font.AddNewColor();
                    clr.indexed = (uint)(value);
                    clr.indexedSpecified = true;
                }
            }
        }


        /**
         *
         * @return xssf color wrapper or null if color info is missing
         */
        public XSSFColor GetXSSFColor()
        {
            if (_font.sizeOfColorArray() == 0) return null;

            return new XSSFColor(_font.GetColorArray(0));
        }

        /**
         * Gets the height of the font in 1/20th point units
         *
         * @return fontheight (in points/20); or -1 if not modified
         */
        public int FontHeight
        {
            get
            {
                if (_font.sizeOfSzArray() == 0) return -1;

                CT_FontSize sz = _font.GetSzArray(0);
                return (short)(20 * sz.val);
            }
            set 
            {
                _font.SetSzArray(null);
                if (value != -1)
                {
                    _font.AddNewSz().val = (double)value / 20;
                }
            }
        }

        /**
         * Get the type of underlining for the font
         *
         * @return font underlining type
         *
         * @see #U_NONE
         * @see #U_SINGLE
         * @see #U_DOUBLE
         * @see #U_SINGLE_ACCOUNTING
         * @see #U_DOUBLE_ACCOUNTING
         */
        public FontUnderlineType UnderlineType
        {
            get
            {
                if (_font.sizeOfUArray() == 0) return FontUnderlineType.None;
                CT_UnderlineProperty u = _font.GetUArray(0);
                switch (u.val)
                {
                    case ST_UnderlineValues.single: return FontUnderlineType.Single;
                    case ST_UnderlineValues.@double: return FontUnderlineType.Double;
                    case ST_UnderlineValues.singleAccounting: return FontUnderlineType.SingleAccounting;
                    case ST_UnderlineValues.doubleAccounting: return FontUnderlineType.DoubleAccounting;
                    default: return FontUnderlineType.None;
                }
            }
            set 
            {
                _font.SetUArray(null);
                if (value != FontUnderlineType.None)
                {
                    FontUnderline fenum = FontUnderline.ValueOf(value);
                    ST_UnderlineValues val = (ST_UnderlineValues)(fenum.Value);
                    _font.AddNewU().val = val;
                }
            }
        }


        /**
         * Get whether the font weight is Set to bold or not
         *
         * @return bold - whether the font is bold or not
         */
        public bool IsBold
        {
            get
            {
                return _font.SizeOfBArray() == 1 && _font.GetBArray(0).val;
            }
        }

        /**
         * @return true if font style was Set to <i>italic</i>
         */
        public bool IsItalic
        {
            get
            {
                return _font.sizeOfIArray() == 1 && _font.GetIArray(0).val;
            }
        }

        /**
         * Set font style options.
         *
         * @param italic - if true, Set posture style to italic, otherwise to normal
         * @param bold if true, Set font weight to bold, otherwise to normal
         */
        public void SetFontStyle(bool italic, bool bold)
        {
            _font.SetIArray(null);
            _font.SetBArray(null);
            if (italic) _font.AddNewI().val = true;
            if (bold) _font.AddNewB().val = true;
        }

        /**
         * Set font style options to default values (non-italic, non-bold)
         */
        public void ResetFontStyle()
        {
            _font = new CT_Font();
        }
    }


}

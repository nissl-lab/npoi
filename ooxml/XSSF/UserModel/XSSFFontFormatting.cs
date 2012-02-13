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
        XSSFFontFormatting(CT_Font font)
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
        public short GetEscapementType()
        {
            if (_font.sizeOfVertAlignArray() == 0) return SS_NONE;

            CT_VerticalAlignFontProperty prop = _font.GetVertAlignArray(0);
            return (short)(prop.val - 1);
        }

        /**
         * Set the escapement type for the font
         *
         * @param escapementType  super or subscript option
         * @see #SS_NONE
         * @see #SS_SUPER
         * @see #SS_SUB
         */
        public void SetEscapementType(short escapementType)
        {
            _font.SetVertAlignArray(null);
            if (escapementType != SS_NONE)
            {
                _font.AddNewVertAlign().SetVal((ST_VerticalAlignRun)(escapementType + 1));
            }
        }

        /**
         * @return font color index
         */
        public short GetFontColorIndex()
        {
            if (_font.sizeOfColorArray() == 0) return -1;

            int idx = 0;
            CT_Color color = _font.GetColorArray(0);
            if (color.IsSetIndexed()) idx = (int)color.indexed;
            return (short)idx;
        }


        /**
         * @param color font color index
         */
        public void SetFontColorIndex(short color)
        {
            _font.SetColorArray(null);
            if (color != -1)
            {
                _font.AddNewColor().indexed = (color);
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
        public int GetFontHeight()
        {
            if (_font.sizeOfSzArray() == 0) return -1;

            CT_FontSize sz = _font.GetSzArray(0);
            return (short)(20 * sz.val);
        }

        /**
         * Sets the height of the font in 1/20th point units
         *
         * @param height the height in twips (in points/20)
         */
        public void SetFontHeight(int height)
        {
            _font.SetSzArray(null);
            if (height != -1)
            {
                _font.AddNewSz().SetVal((double)height / 20);
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
        public short GetUnderlineType()
        {
            if (_font.sizeOfUArray() == 0) return FontUnderline.NONE;
            CT_UnderlineProperty u = _font.GetUArray(0);
            switch (u.val)
            {
                case ST_UnderlineValues.single: return FontUnderline.SINGLE;
                case ST_UnderlineValues.@double: return FontUnderline.DOUBLE;
                case ST_UnderlineValues.singleAccounting: return FontUnderline.SINGLE_ACCOUNTING;
                case ST_UnderlineValues.doubleAccounting: return FontUnderline.DOUBLE_ACCOUNTING;
                default: return FontUnderline.NONE;
            }
        }

        /**
         * Set the type of underlining type for the font
         *
         * @param underlineType  super or subscript option
         *
         * @see #U_NONE
         * @see #U_SINGLE
         * @see #U_DOUBLE
         * @see #U_SINGLE_ACCOUNTING
         * @see #U_DOUBLE_ACCOUNTING
         */
        public void SetUnderlineType(short underlineType)
        {
            _font.SetUArray(null);
            if (underlineType != U_NONE)
            {
                FontUnderline fenum = FontUnderline.ValueOf(underlineType);
                ST_UnderlineValues val = (ST_UnderlineValues)(fenum.Value);
                _font.AddNewU().val =val;
            }
        }

        /**
         * Get whether the font weight is Set to bold or not
         *
         * @return bold - whether the font is bold or not
         */
        public bool IsBold()
        {
            return _font.sizeOfBArray() == 1 && _font.GetBArray(0).val;
        }

        /**
         * @return true if font style was Set to <i>italic</i>
         */
        public bool IsItalic()
        {
            return _font.sizeOfIArray() == 1 && _font.GetIArray(0).GetVal();
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
            if (italic) _font.AddNewI().SetVal(true);
            if (bold) _font.AddNewB().SetVal(true);
        }

        /**
         * Set font style options to default values (non-italic, non-bold)
         */
        public void ResetFontStyle()
        {
            _font.Set(new CT_Font());
        }
    }


}

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

using NPOI.SS.UserModel;
using System;
using Spreadsheet=NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Dml;
using Dml = NPOI.OpenXmlFormats.Dml;
using NPOI.XSSF.Model;
namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a font used in a workbook.
     *
     * @author Gisella Bronzetti
     */
    public class XSSFFont : IFont
    {

        /**
         * By default, Microsoft Office Excel 2007 uses the Calibry font in font size 11
         */
        public const String DEFAULT_FONT_NAME = "Calibri";
        /**
         * By default, Microsoft Office Excel 2007 uses the Calibry font in font size 11
         */
        public const short DEFAULT_FONT_SIZE = 11;
        /**
         * Default font color is black
         * @see NPOI.SS.usermodel.IndexedColors#BLACK
         */
        public static short DEFAULT_FONT_COLOR = IndexedColors.Black.Index;

        private ThemesTable _themes;
        private CT_Font _ctFont;
        private short _index;

        /**
         * Create a new XSSFFont
         *
         * @param font the underlying CT_Font bean
         */
        public XSSFFont(CT_Font font)
        {
            _ctFont = font;
            _index = 0;
        }

        public XSSFFont(CT_Font font, int index)
        {
            _ctFont = font;
            _index = (short)index;
        }

        /**
         * Create a new XSSFont. This method is protected to be used only by XSSFWorkbook
         */
        public XSSFFont()
        {
            this._ctFont = new CT_Font();
            FontName = DEFAULT_FONT_NAME;
            FontHeight =DEFAULT_FONT_SIZE;
        }

        /**
         * get the underlying CT_Font font
         */

        public CT_Font GetCTFont()
        {
            return _ctFont;
        }

        /**
         * get a bool value for the boldness to use.
         *
         * @return bool - bold
         */
        public bool IsBold
        {
            get
            {
                CT_BooleanProperty bold = _ctFont.SizeOfBArray() == 0 ? null : _ctFont.GetBArray(0);
                return (bold != null && bold.val);
            }
            set 
            {
                if (value)
                {
                    CT_BooleanProperty ctBold = _ctFont.SizeOfBArray() == 0 ? _ctFont.AddNewB() : _ctFont.GetBArray(0);
                    ctBold.val = value;
                }
                else
                {
                    _ctFont.SetBArray(null);
                }
            }
        }

        /**
         * get character-set to use.
         *
         * @return int - character-set (0-255)
         * @see NPOI.SS.usermodel.FontCharset
         */
        public short Charset
        {
            get
            {
                CT_IntProperty charset = _ctFont.sizeOfCharsetArray() == 0 ? null : _ctFont.GetCharsetArray(0);
                int val = charset == null ? FontCharset.ANSI.Value : FontCharset.ValueOf(charset.val).Value;
                return (short)val;
            }
            set
            {

            }
        }


        /**
         * get the indexed color value for the font
         * References a color defined in IndexedColors.
         *
         * @return short - indexed color to use
         * @see IndexedColors
         */
        public short Color
        {
            get
            {
                Spreadsheet.CT_Color color = _ctFont.sizeOfColorArray() == 0 ? null : _ctFont.GetColorArray(0);
                if (color == null) return IndexedColors.Black.Index;

                if (!color.indexedSpecified) return IndexedColors.Black.Index;
                long index = color.indexed;
                if (index == XSSFFont.DEFAULT_FONT_COLOR)
                {
                    return IndexedColors.Black.Index;
                }
                else if (index == IndexedColors.Red.Index)
                {
                    return IndexedColors.Red.Index;
                }
                else
                {
                    return (short)index;
                }
            }
            set 
            {
                Spreadsheet.CT_Color ctColor = _ctFont.sizeOfColorArray() == 0 ? _ctFont.AddNewColor() : _ctFont.GetColorArray(0);
                switch (value)
                {
                    case (short)FontColor.Normal:
                        
                            ctColor.indexed = (uint)(XSSFFont.DEFAULT_FONT_COLOR);
                            ctColor.indexedSpecified = true;
                            break;
                        
                    case (short)FontColor.Red:

                            ctColor.indexed = (uint)(IndexedColors.Red.Index);
                            ctColor.indexedSpecified = true;
                            break;
                        
                    default:
                            ctColor.indexed = (uint)(value);
                            ctColor.indexedSpecified = true;
                        break;
                }
            }
        }


        /**
         * get the color value for the font
         * References a color defined as  Standard Alpha Red Green Blue color value (ARGB).
         *
         * @return XSSFColor - rgb color to use
         */
        public XSSFColor GetXSSFColor()
        {
            Spreadsheet.CT_Color ctColor = _ctFont.sizeOfColorArray() == 0 ? null : _ctFont.GetColorArray(0);
            if (ctColor != null)
            {
                XSSFColor color = new XSSFColor(ctColor);
                if (_themes != null)
                {
                    _themes.InheritFromThemeAsRequired(color);
                }
                return color;
            }
            else
            {
                return null;
            }
        }


        /**
         * get the color value for the font
         * References a color defined in theme.
         *
         * @return short - theme defined to use
         */
        public short GetThemeColor()
        {
            Spreadsheet.CT_Color color = _ctFont.sizeOfColorArray() == 0 ? null : _ctFont.GetColorArray(0);
            long index = ((color == null) || !color.themeSpecified) ? 0 : color.theme;
            return (short)index;
        }

        /**
         * get the font height in point.
         *
         * @return short - height in point
         */
        public double FontHeight
        {
            get
            {
                CT_FontSize size = _ctFont.sizeOfSzArray() == 0 ? null : _ctFont.GetSzArray(0);
                if (size != null)
                {
                    double fontHeight = size.val;
                    return (short)(fontHeight * 20);
                }
                return (short)(DEFAULT_FONT_SIZE * 20);
            }
            set 
            {
                CT_FontSize fontSize = _ctFont.sizeOfSzArray() == 0 ? _ctFont.AddNewSz() : _ctFont.GetSzArray(0);
                fontSize.val = value;
            }
        }


        /**
         * @see #GetFontHeight()
         */
        public short FontHeightInPoints
        {
            get
            {
                return (short)(FontHeight / 20);
            }
            set 
            {
                CT_FontSize fontSize = _ctFont.sizeOfSzArray() == 0 ? _ctFont.AddNewSz() : _ctFont.GetSzArray(0);
                fontSize.val = value;
            }
        }

        /**
         * get the name of the font (i.e. Arial)
         *
         * @return String - a string representing the name of the font to use
         */
        public String FontName
        {
            get
            {
                CT_FontName name = _ctFont.name;
                return name == null ? DEFAULT_FONT_NAME : name.val;
            }
            set 
            {
                CT_FontName fontName = _ctFont.name==null?_ctFont.AddNewName():_ctFont.name;
                fontName.val = value == null ? DEFAULT_FONT_NAME : value;
            }
        }

        /**
         * get a bool value that specify whether to use italics or not
         *
         * @return bool - value for italic
         */
        public bool IsItalic
        {
            get
            {
                CT_BooleanProperty italic = _ctFont.sizeOfIArray() == 0 ? null : _ctFont.GetIArray(0);
                return italic != null && italic.val;
            }
            set 
            {
                if (value)
                {
                    CT_BooleanProperty bool1 = _ctFont.sizeOfIArray() == 0 ? _ctFont.AddNewI() : _ctFont.GetIArray(0);
                    bool1.val = value;
                }
                else
                {
                    _ctFont.SetIArray(null);
                }
            }
        }

        /**
         * get a bool value that specify whether to use a strikeout horizontal line through the text or not
         *
         * @return bool - value for strikeout
         */
        public bool IsStrikeout
        {
            get
            {
                CT_BooleanProperty strike = _ctFont.sizeOfStrikeArray() == 0 ? null : _ctFont.GetStrikeArray(0);
                return strike != null && strike.val;
            }
            set 
            {
                if (!value) _ctFont.SetStrikeArray(null);
                else
                {
                    CT_BooleanProperty strike = _ctFont.sizeOfStrikeArray() == 0 ? _ctFont.AddNewStrike() : _ctFont.GetStrikeArray(0);
                    strike.val = value;
                }
            }
        }

        /**
         * get normal,super or subscript.
         *
         * @return short - offset type to use (none,super,sub)
         * @see Font#SS_NONE
         * @see Font#SS_SUPER
         * @see Font#SS_SUB
         */
        public FontSuperScript TypeOffset
        {
            get
            {
                CT_VerticalAlignFontProperty vAlign = _ctFont.sizeOfVertAlignArray() == 0 ? null : _ctFont.GetVertAlignArray(0);
                if (vAlign == null)
                {
                    return FontSuperScript.None;
                }
                ST_VerticalAlignRun val = vAlign.val;
                switch (val)
                {
                    case ST_VerticalAlignRun.baseline:
                        return FontSuperScript.None;
                    case ST_VerticalAlignRun.subscript:
                        return FontSuperScript.Sub;
                    case ST_VerticalAlignRun.superscript:
                        return FontSuperScript.Super;
                    default:
                        throw new POIXMLException("Wrong offset value " + val);
                }
            }
            set 
            {
                if (value == (short)FontSuperScript.None)
                {
                    _ctFont.SetVertAlignArray(null);
                }
                else
                {
                    CT_VerticalAlignFontProperty offSetProperty = _ctFont.sizeOfVertAlignArray() == 0 ? _ctFont.AddNewVertAlign() : _ctFont.GetVertAlignArray(0);
                    switch (value)
                    {
                        case FontSuperScript.None:
                            offSetProperty.val = ST_VerticalAlignRun.baseline;
                            break;
                        case FontSuperScript.Sub:
                            offSetProperty.val = ST_VerticalAlignRun.subscript;
                            break;
                        case FontSuperScript.Super:
                            offSetProperty.val = ST_VerticalAlignRun.superscript;
                            break;
                    }
                }
            }
        }

        /**
         * get type of text underlining to use
         *
         * @return byte - underlining type
         * @see NPOI.SS.usermodel.FontUnderline
         */
        public FontUnderlineType Underline
        {
            get
            {
                CT_UnderlineProperty underline = _ctFont.sizeOfUArray() == 0 ? null : _ctFont.GetUArray(0);
                if (underline != null)
                {
                    FontUnderline val = FontUnderline.ValueOf((int)underline.val);
                    return (FontUnderlineType)val.ByteValue;
                }
                return (FontUnderlineType)FontUnderline.NONE.ByteValue;
            }
            set 
            {
                SetUnderline(value);
            }
        }

        /**
         * get the boldness to use
         * @return boldweight
         * @see #BOLDWEIGHT_NORMAL
         * @see #BOLDWEIGHT_BOLD
         */

        public short Boldweight
        {
            get
            {
                return (IsBold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal);
            }
            set 
            {
                this.IsBold = (value == (short)FontBoldWeight.Bold);
            }
        }

        /**
         * set character-set to use.
         *
         * @param charset - charset
         * @see FontCharset
         */
        public void SetCharSet(byte charSet)
        {
            int cs = (int)charSet;
            if (cs < 0)
            {
                cs += 256;
            }
            SetCharSet(cs);
        }
        /**
         * set character-set to use.
         *
         * @param charset - charset
         * @see FontCharset
         */
        public void SetCharSet(int charset)
        {
            FontCharset FontCharset = FontCharset.ValueOf(charset);
            if (FontCharset != null)
            {
                SetCharSet(FontCharset);
            }
            else
            {
                throw new POIXMLException("Attention: an attempt to set a type of unknow charset and charSet");
            }
        }

        /**
         * set character-set to use.
         *
         * @param charSet
         */
        public void SetCharSet(FontCharset charset)
        {
            CT_IntProperty charSetProperty;
            if (_ctFont.sizeOfCharsetArray() == 0)
            {
                charSetProperty = _ctFont.AddNewCharset();
            }
            else
            {
                charSetProperty = _ctFont.GetCharsetArray(0);
            }
            // We know that FontCharset only has valid entries in it,
            //  so we can just set the int value from it
            charSetProperty.val = charset.Value;
        }

        /**
         * set the color for the font in Standard Alpha Red Green Blue color value
         *
         * @param color - color to use
         */
        public void SetColor(XSSFColor color)
        {
            if (color == null) _ctFont.SetColorArray(null);
            else
            {
                Spreadsheet.CT_Color ctColor = _ctFont.sizeOfColorArray() == 0 ? _ctFont.AddNewColor() : _ctFont.GetColorArray(0);
                ctColor.SetRgb(color.RGB);
            }
        }


        /**
         * set the theme color for the font to use
         *
         * @param theme - theme color to use
         */
        public void SetThemeColor(short theme)
        {
            Spreadsheet.CT_Color ctColor = _ctFont.sizeOfColorArray() == 0 ? _ctFont.AddNewColor() : _ctFont.GetColorArray(0);
            ctColor.theme = (uint)theme;
        }




        /**
         * set an enumeration representing the style of underlining that is used.
         * The none style is equivalent to not using underlining at all.
         * The possible values for this attribute are defined by the FontUnderline
         *
         * @param underline - FontUnderline enum value
         */
        internal void SetUnderline(FontUnderlineType underline)
        {
            if (underline == FontUnderlineType.None)
            {
                _ctFont.SetUArray(null);
            }
            else
            {
                CT_UnderlineProperty ctUnderline = _ctFont.sizeOfUArray() == 0 ? _ctFont.AddNewU() : _ctFont.GetUArray(0);
                ST_UnderlineValues val = (ST_UnderlineValues)FontUnderline.ValueOf(underline).Value;
                ctUnderline.val = val;
            }
        }


        public override String ToString()
        {
            return _ctFont.ToString();
        }


        ///**
        // * Perform a registration of ourselves 
        // *  to the style table
        // */
        public long RegisterTo(StylesTable styles)
        {
            this._themes = styles.GetTheme();
            short idx = (short)styles.PutFont(this, true);
            this._index = idx;
            return idx;
        }
        /**
         * Records the Themes Table that is associated with
         *  the current font, used when looking up theme
         *  based colours and properties.
         */
        public void SetThemesTable(ThemesTable themes)
        {
            this._themes = themes;
        }

        /**
         * get the font scheme property.
         * is used only in StylesTable to create the default instance of font
         *
         * @return FontScheme
         * @see NPOI.XSSF.model.StylesTable#CreateDefaultFont()
         */
        public FontScheme GetScheme()
        {
            NPOI.OpenXmlFormats.Spreadsheet.CT_FontScheme scheme = _ctFont.sizeOfSchemeArray() == 0 ? null : _ctFont.GetSchemeArray(0);
            return scheme == null ? FontScheme.NONE : FontScheme.ValueOf((int)scheme.val);
        }

        /**
         * set font scheme property
         *
         * @param scheme - FontScheme enum value
         * @see FontScheme
         */
        public void SetScheme(FontScheme scheme)
        {
            NPOI.OpenXmlFormats.Spreadsheet.CT_FontScheme ctFontScheme = _ctFont.sizeOfSchemeArray() == 0 ? _ctFont.AddNewScheme() : _ctFont.GetSchemeArray(0);
            ST_FontScheme val = (ST_FontScheme)scheme.Value;
            ctFontScheme.val = val;
        }

        /**
         * get the font family to use.
         *
         * @return the font family to use
         * @see NPOI.SS.usermodel.FontFamily
         */
        public int Family
        {
            get
            {
                CT_IntProperty family = _ctFont.sizeOfFamilyArray() == 0 ? _ctFont.AddNewFamily() : _ctFont.GetFamilyArray(0);
                return family == null ? FontFamily.NOT_APPLICABLE.Value : FontFamily.ValueOf(family.val).Value;
            }
            set 
            {
                CT_IntProperty family = _ctFont.sizeOfFamilyArray() == 0 ? _ctFont.AddNewFamily() : _ctFont.GetFamilyArray(0);
                family.val = value;
            }
        }


        /**
         * set an enumeration representing the font family this font belongs to.
         * A font family is a set of fonts having common stroke width and serif characteristics.
         *
         * @param family font family
         * @link #SetFamily(int value)
         */
        public void SetFamily(FontFamily family)
        {
            Family = family.Value;
        }

        /**
         * get the index within the XSSFWorkbook (sequence within the collection of Font objects)
         * @return unique index number of the underlying record this Font represents (probably you don't care
         *  unless you're comparing which one is which)
         */
        public short Index
        {
            get
            {
                return _index;
            }
        }

        public override int GetHashCode()
        {
            return _ctFont.ToString().GetHashCode();
        }

        public override bool Equals(Object o)
        {
            if (!(o is XSSFFont)) return false;

            XSSFFont cf = (XSSFFont)o;
            return _ctFont.ToString().Equals(cf.GetCTFont().ToString());
        }

    }
}

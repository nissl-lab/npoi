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
using System;
using NPOI.OpenXmlFormats.Dml;
using System.Drawing;
using NPOI.Util;
namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a run of text within the Containing text body. The run element is the
     * lowest level text separation mechanism within a text body.
     */
    public class XSSFTextRun
    {
        private CT_RegularTextRun _r;
        private XSSFTextParagraph _p;

        public XSSFTextRun(CT_RegularTextRun r, XSSFTextParagraph p)
        {
            _r = r;
            _p = p;
        }

        XSSFTextParagraph GetParentParagraph()
        {
            return _p;
        }

        public String GetText()
        {
            return _r.t;
        }

        public void SetText(String text)
        {
            _r.t = (text);
        }

        public CT_RegularTextRun GetXmlObject()
        {
            return _r;
        }

        public void SetFontColor(Color color)
        {
            CT_TextCharacterProperties rPr = GetRPr();
            CT_SolidColorFillProperties fill = rPr.IsSetSolidFill() ? rPr.solidFill : rPr.AddNewSolidFill();
            CT_SRgbColor clr = fill.IsSetSrgbClr() ? fill.srgbClr : fill.AddNewSrgbClr();
            clr.val = (new byte[] { color.R, color.G, color.B });
            
            if (fill.IsSetHslClr()) fill.UnsetHslClr();
            if (fill.IsSetPrstClr()) fill.UnsetPrstClr();
            if (fill.IsSetSchemeClr()) fill.UnsetSchemeClr();
            if (fill.IsSetScrgbClr()) fill.UnsetScrgbClr();
            if (fill.IsSetSysClr()) fill.UnsetSysClr();

        }

        public Color GetFontColor()
        {

            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetSolidFill())
            {
                CT_SolidColorFillProperties fill = rPr.solidFill;

                if (fill.IsSetSrgbClr())
                {
                    CT_SRgbColor clr = fill.srgbClr;
                    byte[] rgb = clr.val;
                    return Color.FromArgb(0xFF & rgb[0], 0xFF & rgb[1], 0xFF & rgb[2]);
                }
            }

            return  Color.FromArgb(0, 0, 0);
        }

        /**
         *
         * @param fontSize  font size in points.
         * The value of <code>-1</code> unsets the Sz attribute from the underlying xml bean
         */
        public void SetFontSize(double fontSize)
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (fontSize == -1.0)
            {
                if (rPr.IsSetSz()) rPr.UnsetSz();
            }
            else
            {
                if (fontSize < 1.0)
                {
                    throw new ArgumentException("Minimum font size is 1pt but was " + fontSize);
                }

                rPr.sz = ((int)(100 * fontSize));
            }
        }

        /**
         * @return font size in points or -1 if font size is not Set.
         */
        public double GetFontSize()
        {
            double scale = 1;
            double size = XSSFFont.DEFAULT_FONT_SIZE;	// default font size
            CT_TextNormalAutofit afit = GetParentParagraph().GetParentShape().txBody.bodyPr.normAutofit;
            if (afit != null) scale = (double)afit.fontScale / 100000;

            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetSz())
            {
                size = rPr.sz * 0.01;
            }

            return size * scale;
        }

        /**
         *
         * @return the spacing between characters within a text Run,
         * If this attribute is omitted then a value of 0 or no adjustment is assumed.
         */
        public double GetCharacterSpacing()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetSpc())
            {
                return rPr.spc * 0.01;
            }
            return 0;
        }

        /**
         * Set the spacing between characters within a text Run.
         * <p>
         * The spacing is specified in points. Positive values will cause the text to expand,
         * negative values to condense.
         * </p>
         *
         * @param spc  character spacing in points.
         */
        public void SetCharacterSpacing(double spc)
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (spc == 0.0)
            {
                if (rPr.IsSetSpc()) rPr.UnsetSpc();
            }
            else
            {
                rPr.spc = ((int)(100 * spc));
            }
        }

        /**
         * Specifies the typeface, or name of the font that is to be used for this text Run.
         *
         * @param typeface  the font to apply to this text Run.
         * The value of <code>null</code> unsets the Typeface attribute from the underlying xml.
         */
        public void SetFont(String typeface)
        {
            SetFontFamily(typeface, unchecked((byte)-1), unchecked((byte)-1), false);
        }

        public void SetFontFamily(String typeface, byte charset, byte pictAndFamily, bool isSymbol)
        {
            CT_TextCharacterProperties rPr = GetRPr();

            if (typeface == null)
            {
                if (rPr.IsSetLatin()) rPr.UnsetLatin();
                if (rPr.IsSetCs()) rPr.UnsetCs();
                if (rPr.IsSetSym()) rPr.UnsetSym();
            }
            else
            {
                if (isSymbol)
                {
                    CT_TextFont font = rPr.IsSetSym() ? rPr.sym : rPr.AddNewSym();
                    font.typeface = (typeface);
                }
                else
                {
                    CT_TextFont latin = rPr.IsSetLatin() ? rPr.latin : rPr.AddNewLatin();
                    latin.typeface = (typeface);
                    if ((sbyte)charset != -1) latin.charset = (sbyte)(charset);
                    if ((sbyte)pictAndFamily != -1) latin.pitchFamily = (sbyte)(pictAndFamily);
                }
            }
        }

        /**
         * @return  font family or null if not Set
         */
        public String GetFontFamily()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            CT_TextFont font = rPr.latin;
            if (font != null)
            {
                return font.typeface;
            }
            return XSSFFont.DEFAULT_FONT_NAME;
        }

        public byte GetPitchAndFamily()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            CT_TextFont font = rPr.latin;
            if (font != null)
            {
                return (byte)font.pitchFamily;
            }
            return 0;
        }

        /**
         * Specifies whether a run of text will be formatted as strikethrough text.
         *
         * @param strike whether a run of text will be formatted as strikethrough text.
         */
        public void SetStrikethrough(bool strike)
        {
            GetRPr().strike = (strike ? ST_TextStrikeType.sngStrike : ST_TextStrikeType.noStrike);
        }

        /**
         * @return whether a run of text will be formatted as strikethrough text. Default is false.
         */
        public bool IsStrikethrough()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetStrike())
            {
                return rPr.strike != ST_TextStrikeType.noStrike;
            }
            return false;
        }

        /**
         * @return whether a run of text will be formatted as a superscript text. Default is false.
         */
        public bool IsSuperscript()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetBaseline())
            {
                return rPr.baseline> 0;
            }
            return false;
        }

        /**
         *  Set the baseline for both the superscript and subscript fonts.
         *  <p>
         *     The size is specified using a percentage.
         *     Positive values indicate superscript, negative values indicate subscript.
         *  </p>
         *
         * @param baselineOffset
         */
        public void SetBaselineOffset(double baselineOffset)
        {
            GetRPr().baseline = ((int)baselineOffset * 1000);
        }

        /**
         * Set whether the text in this run is formatted as superscript.
         * Default base line offset is 30%
         *
         * @see #setBaselineOffset(double)
         */
        public void SetSuperscript(bool flag){
        SetBaselineOffset(flag ? 30.0 : 0.0);
    }

        /**
         * Set whether the text in this run is formatted as subscript.
         * Default base line offset is -25%.
         *
         * @see #setBaselineOffset(double)
         */
        public void SetSubscript(bool flag){
        SetBaselineOffset(flag ? -25.0 : 0.0);
    }

        /**
         * @return whether a run of text will be formatted as a superscript text. Default is false.
         */
        public bool IsSubscript()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetBaseline())
            {
                return rPr.baseline < 0;
            }
            return false;
        }

        /**
         * @return whether a run of text will be formatted as a superscript text. Default is false.
         */
        public TextCap GetTextCap()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetCap())
            {
                return EnumConverter.ValueOf<TextCap, ST_TextCapsType>(rPr.cap);
            }
            return TextCap.NONE;
        }

        /**
         * Specifies whether this run of text will be formatted as bold text
         *
         * @param bold whether this run of text will be formatted as bold text
         */
        public void SetBold(bool bold)
        {
            GetRPr().b = (bold);
        }

        /**
         * @return whether this run of text is formatted as bold text
         */
        public bool IsBold()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetB())
            {
                return rPr.b;
            }
            return false;
        }

        /**
         * @param italic whether this run of text is formatted as italic text
         */
        public void SetItalic(bool italic)
        {
            GetRPr().i = (italic);
        }

        /**
         * @return whether this run of text is formatted as italic text
         */
        public bool IsItalic()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetI())
            {
                return rPr.i;
            }
            return false;
        }

        /**
         * @param underline whether this run of text is formatted as underlined text
         */
        public void SetUnderline(bool underline)
        {
            GetRPr().u = (underline ? ST_TextUnderlineType.sng : ST_TextUnderlineType.none);
        }

        /**
         * @return whether this run of text is formatted as underlined text
         */
        public bool IsUnderline()
        {
            CT_TextCharacterProperties rPr = GetRPr();
            if (rPr.IsSetU())
            {
                return rPr.u != ST_TextUnderlineType.none;
            }
            return false;
        }

        internal CT_TextCharacterProperties GetRPr()
        {
            return _r.IsSetRPr() ? _r.rPr : _r.AddNewRPr();
        }


        public String ToString()
        {
            return "[" + this.GetType().ToString() + "]" + GetText();
        }
    }

}
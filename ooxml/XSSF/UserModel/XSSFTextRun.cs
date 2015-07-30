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

        public XSSFTextParagraph ParentParagraph
        {
            get
            {
                return _p;
            }
        }

        public String Text
        {
            get
            {
                return _r.t;
            }
            set
            {
                _r.t = (value);
            }
        }

        public CT_RegularTextRun GetXmlObject()
        {
            return _r;
        }

        public Color FontColor
        {
            get
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

                return Color.FromArgb(0, 0, 0);
            }
            set
            {
                CT_TextCharacterProperties rPr = GetRPr();
                CT_SolidColorFillProperties fill = rPr.IsSetSolidFill() ? rPr.solidFill : rPr.AddNewSolidFill();
                CT_SRgbColor clr = fill.IsSetSrgbClr() ? fill.srgbClr : fill.AddNewSrgbClr();
                clr.val = (new byte[] { value.R, value.G, value.B });

                if (fill.IsSetHslClr()) fill.UnsetHslClr();
                if (fill.IsSetPrstClr()) fill.UnsetPrstClr();
                if (fill.IsSetSchemeClr()) fill.UnsetSchemeClr();
                if (fill.IsSetScrgbClr()) fill.UnsetScrgbClr();
                if (fill.IsSetSysClr()) fill.UnsetSysClr();
            }
        }

        /**
         * @return font size in points or -1 if font size is not Set.
         */
        public double FontSize
        {
            get
            {
                double scale = 1;
                double size = XSSFFont.DEFAULT_FONT_SIZE;	// default font size
                CT_TextNormalAutofit afit = ParentParagraph.ParentShape.txBody.bodyPr.normAutofit;
                if (afit != null) scale = (double)afit.fontScale / 100000;

                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetSz())
                {
                    size = rPr.sz * 0.01;
                }

                return size * scale;
            }
            set
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (value == -1.0)
                {
                    if (rPr.IsSetSz()) rPr.UnsetSz();
                }
                else
                {
                    if (value < 1.0)
                    {
                        throw new ArgumentException("Minimum font size is 1pt but was " + value);
                    }

                    rPr.sz = ((int)(100 * value));
                }
            }
        }

        /**
         *
         * @return the spacing between characters within a text Run,
         * If this attribute is omitted then a value of 0 or no adjustment is assumed.
         */
        public double CharacterSpacing
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetSpc())
                {
                    return rPr.spc * 0.01;
                }
                return 0;
            }
            set
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (value == 0.0)
                {
                    if (rPr.IsSetSpc()) rPr.UnsetSpc();
                }
                else
                {
                    rPr.spc = ((int)(100 * value));
                }
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
        public String FontFamily
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                CT_TextFont font = rPr.latin;
                if (font != null)
                {
                    return font.typeface;
                }
                return XSSFFont.DEFAULT_FONT_NAME;
            }
        }

        public byte PitchAndFamily
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                CT_TextFont font = rPr.latin;
                if (font != null)
                {
                    return (byte)font.pitchFamily;
                }
                return 0;
            }
        }

        /**
         * get or set whether a run of text will be formatted as strikethrough text. Default is false.
         */
        public bool IsStrikethrough
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetStrike())
                {
                    return rPr.strike != ST_TextStrikeType.noStrike;
                }
                return false;
            }
            set
            {
                GetRPr().strike = (value ? ST_TextStrikeType.sngStrike : ST_TextStrikeType.noStrike);
            }
        }

        /**
         * get or set whether a run of text will be formatted as a superscript text. Default is false.
         * Default base line offset is 30%
         */
        public bool IsSuperscript
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetBaseline())
                {
                    return rPr.baseline > 0;
                }
                return false;
            }
            set
            {
                SetBaselineOffset(value ? 30.0 : 0.0);
            }
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
         * get or set whether a run of text will be formatted as a superscript text. Default is false.
         * Default base line offset is -25%.
         */
        public bool IsSubscript
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetBaseline())
                {
                    return rPr.baseline < 0;
                }
                return false;
            }
            set
            {
                SetBaselineOffset(value ? -25.0 : 0.0);
            }
        }

        /**
         * @return whether a run of text will be formatted as a superscript text. Default is false.
         */
        public TextCap TextCap
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetCap())
                {
                    return EnumConverter.ValueOf<TextCap, ST_TextCapsType>(rPr.cap);
                }
                return TextCap.NONE;
            }
        }

        /**
         * get or set whether this run of text is formatted as bold text
         */
        public bool IsBold
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetB())
                {
                    return rPr.b;
                }
                return false;
            }
            set
            {
                GetRPr().b = value;
            }
        }

        /**
         * get or set whether this run of text is formatted as italic text
         */
        public bool IsItalic
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetI())
                {
                    return rPr.i;
                }
                return false;
            }
            set
            {
                GetRPr().i = value;
            }
        }

        /**
         * get or set whether this run of text is formatted as underlined text
         */
        public bool IsUnderline
        {
            get
            {
                CT_TextCharacterProperties rPr = GetRPr();
                if (rPr.IsSetU())
                {
                    return rPr.u != ST_TextUnderlineType.none;
                }
                return false;
            }
            set
            {
                GetRPr().u = (value ? ST_TextUnderlineType.sng : ST_TextUnderlineType.none);
            }
        }

        internal CT_TextCharacterProperties GetRPr()
        {
            return _r.IsSetRPr() ? _r.rPr : _r.AddNewRPr();
        }


        public override String ToString()
        {
            return "[" + this.GetType().ToString() + "]" + Text;
        }
    }

}
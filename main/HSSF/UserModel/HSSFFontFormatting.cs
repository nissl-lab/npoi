/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using System;

    /// <summary>
    /// High level representation for Font Formatting component
    /// of Conditional Formatting Settings
    /// </summary>
    ///
    /// @author Dmitriy Kumshayev
    public class HSSFFontFormatting : IFontFormatting
    {

        private FontFormatting fontFormatting;
        private HSSFWorkbook workbook;
        public HSSFFontFormatting(CFRuleBase cfRuleRecord, HSSFWorkbook workbook)
        {
            this.fontFormatting = cfRuleRecord.FontFormatting;
            this.workbook = workbook;
        }

        protected FontFormatting GetFontFormattingBlock()
        {
            return fontFormatting;
        }

        /// <summary>
        /// Get the type of base or subscript for the font
        /// </summary>
        /// <returns>base or subscript option</returns>
        public FontSuperScript EscapementType
        {
            get
            {
                return (FontSuperScript) fontFormatting.EscapementType;
            }
            set
            {
                switch(value)
                {
                    case FontSuperScript.Sub:
                    case FontSuperScript.Super:
                        fontFormatting.EscapementType = value;
                        fontFormatting.IsEscapementTypeModified = true;
                        break;
                    case FontSuperScript.None:
                        fontFormatting.EscapementType = value;
                        fontFormatting.IsEscapementTypeModified = false;
                        break;
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>font color index</returns>
        public short FontColorIndex
        {
            get
            {
                return fontFormatting.FontColorIndex;
            }
            set { fontFormatting.FontColorIndex=(value); }
        }

        public IColor FontColor
        {
            get
            {
                return workbook.GetCustomPalette().GetColor(FontColorIndex);
            }
            set
            {
                HSSFColor hcolor = HSSFColor.ToHSSFColor(value);
                if(hcolor == null)
                {
                    fontFormatting.FontColorIndex = ((short) 0);
                }
                else
                {
                    fontFormatting.FontColorIndex = (hcolor.Indexed);
                }
            }
        }

        /// <summary>
        /// Gets the height of the font in 1/20th point Units
        /// </summary>
        /// <returns>fontheight (in points/20); or -1 if not modified</returns>
        public int FontHeight
        {
            get { return fontFormatting.FontHeight; }
            set { fontFormatting.FontHeight=(value); }
        }

        /// <summary>
        /// Get the font weight for this font (100-1000dec or 0x64-0x3e8).  Default Is
        /// 0x190 for normal and 0x2bc for bold
        /// </summary>
        /// <returns>bw - a number between 100-1000 for the fonts "boldness"</returns>

        public short FontWeight
        {
            get { return fontFormatting.FontWeight; }
        }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        /// <see cref="org.apache.poi.hssf.record.cf.FontFormatting.GetRawRecord()" />
        protected byte[] GetRawRecord()
        {
            return fontFormatting.RawRecord;
        }

        /// <summary>
        /// Get the type of Underlining for the font
        /// </summary>
        /// <returns>font Underlining type</returns>
        /// 
        /// <see cref="U_NONE" />
        /// <see cref="U_SINGLE" />
        /// <see cref="U_DOUBLE" />
        /// <see cref="U_SINGLE_ACCOUNTING" />
        /// <see cref="U_DOUBLE_ACCOUNTING" />
        public FontUnderlineType UnderlineType
        {
            get
            {
                return (FontUnderlineType) fontFormatting.UnderlineType;
            }
            set
            {
                switch(value)
                {
                    case FontUnderlineType.Single:
                    case FontUnderlineType.Double:
                    case FontUnderlineType.SingleAccounting:
                    case FontUnderlineType.DoubleAccounting:
                        fontFormatting.UnderlineType = value;
                        IsUnderlineTypeModified = true;
                        break;

                    case FontUnderlineType.None:
                        fontFormatting.UnderlineType = value;
                        IsUnderlineTypeModified = false;
                        break;

                }
            }
        }

        /// <summary>
        /// Get whether the font weight Is Set to bold or not
        /// </summary>
        /// <returns>bold - whether the font Is bold or not</returns>
        public bool IsBold
        {
            get
            {
                return fontFormatting.IsFontWeightModified && fontFormatting.IsBold;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if escapement type was modified from default</returns>
        public bool IsEscapementTypeModified
        {
            get
            {
                return fontFormatting.IsEscapementTypeModified;
            }
            set { fontFormatting.IsEscapementTypeModified=value; }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font cancellation was modified from default</returns>
        public bool IsFontCancellationModified
        {
            get
            {
                return fontFormatting.IsFontCancellationModified;
            }
            set { fontFormatting.IsFontCancellationModified=(value); }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font outline type was modified from default</returns>
        public bool IsFontOutlineModified
        {
            get
            {
                return fontFormatting.IsFontOutlineModified;
            }
            set { fontFormatting.IsFontOutlineModified=(value); }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font shadow type was modified from default</returns>
        public bool IsFontShadowModified
        {
            get
            {
                return fontFormatting.IsFontShadowModified;
            }
            set { fontFormatting.IsFontShadowModified=value; }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font style was modified from default</returns>
        public bool IsFontStyleModified
        {
            get
            {
                return fontFormatting.IsFontStyleModified;
            }
            set { fontFormatting.IsFontStyleModified=value; }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font style was Set to <i>italic</i></returns>
        public bool IsItalic
        {
            get { return fontFormatting.IsFontStyleModified && fontFormatting.IsItalic; }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font outline Is on</returns>
        public bool IsOutlineOn
        {
            get
            {
                return fontFormatting.IsFontOutlineModified && fontFormatting.IsOutlineOn;
            }
            set
            {
                fontFormatting.IsOutlineOn=value;
                fontFormatting.IsFontOutlineModified=value;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font shadow Is on</returns>
        public bool IsShadowOn
        {
            get { return fontFormatting.IsFontOutlineModified && fontFormatting.IsShadowOn; }
            set
            {
                fontFormatting.IsShadowOn=value;
                fontFormatting.IsFontShadowModified=value;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font strikeout Is on</returns>
        public bool IsStrikeout
        {
            get
            {
                return fontFormatting.IsFontCancellationModified && fontFormatting.IsStruckout;

            }
            set
            {
                fontFormatting.IsStruckout = (value);
                fontFormatting.IsFontCancellationModified = (value);
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font Underline type was modified from default</returns>
        public bool IsUnderlineTypeModified
        {
            get { return fontFormatting.IsUnderlineTypeModified; }
            set { fontFormatting.IsUnderlineTypeModified=value; }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if font weight was modified from default</returns>
        public bool IsFontWeightModified
        {
            get
            {
                return fontFormatting.IsFontWeightModified;
            }

        }

        /// <summary>
        /// Set font style options.
        /// </summary>
        /// <param name="italic">- if true, Set posture style to italic, otherwise to normal</param>
        /// <param name="bold-">if true, Set font weight to bold, otherwise to normal</param>

        public void SetFontStyle(bool italic, bool bold)
        {
            bool modified = italic || bold;
            fontFormatting.IsItalic=italic;
            fontFormatting.IsBold=bold;
            fontFormatting.IsFontStyleModified=modified;
            fontFormatting.IsFontWeightModified=modified;
        }

        /// <summary>
        /// Set font style options to default values (non-italic, non-bold)
        /// </summary>
        public void ResetFontStyle()
        {
            SetFontStyle(false, false);
        }





    }
}
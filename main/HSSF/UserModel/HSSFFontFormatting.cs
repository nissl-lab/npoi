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
    using NPOI.SS.UserModel;

    /**
     * High level representation for Font Formatting component
     * of Conditional Formatting Settings
     * 
     * @author Dmitriy Kumshayev
     *
     */
    public class HSSFFontFormatting : IFontFormatting
    {

        private FontFormatting fontFormatting;

        public HSSFFontFormatting(CFRuleRecord cfRuleRecord)
        {
            this.fontFormatting = cfRuleRecord.FontFormatting;
        }

        protected FontFormatting GetFontFormattingBlock()
        {
            return fontFormatting;
        }

        /**
         * Get the type of base or subscript for the font
         *
         * @return base or subscript option
         */
        public FontSuperScript EscapementType
        {
            get
            {
                return (FontSuperScript)fontFormatting.EscapementType;
            }
            set
            {
                switch (value)
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

        /**
         * @return font color index
         */
        public short FontColorIndex
        {
            get
            {
                return fontFormatting.FontColorIndex;
            }
            set { fontFormatting.FontColorIndex=(value); }
        }

        /**
         * Gets the height of the font in 1/20th point Units
         *
         * @return fontheight (in points/20); or -1 if not modified
         */
        public int FontHeight
        {
            get { return fontFormatting.FontHeight; }
            set { fontFormatting.FontHeight=(value); }
        }

        /**
         * Get the font weight for this font (100-1000dec or 0x64-0x3e8).  Default Is
         * 0x190 for normal and 0x2bc for bold
         *
         * @return bw - a number between 100-1000 for the fonts "boldness"
         */

        public short FontWeight
        {
            get { return fontFormatting.FontWeight; }
        }

        /**
         * @return
         * @see org.apache.poi.hssf.record.cf.FontFormatting#GetRawRecord()
         */
        protected byte[] GetRawRecord()
        {
            return fontFormatting.GetRawRecord();
        }

        /**
         * Get the type of Underlining for the font
         *
         * @return font Underlining type
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
                return (FontUnderlineType)fontFormatting.UnderlineType;
            }
            set
            {
                switch (value)
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

        /**
         * Get whether the font weight Is Set to bold or not
         *
         * @return bold - whether the font Is bold or not
         */
        public bool IsBold
        {
            get
            {
                return fontFormatting.IsFontWeightModified && fontFormatting.IsBold;
            }
        }

        /**
         * @return true if escapement type was modified from default   
         */
        public bool IsEscapementTypeModified
        {
            get{
                return fontFormatting.IsEscapementTypeModified;
            }
            set { fontFormatting.IsEscapementTypeModified=value; }
        }

        /**
         * @return true if font cancellation was modified from default   
         */
        public bool IsFontCancellationModified
        {
            get{
            return fontFormatting.IsFontCancellationModified;
            }
            set { fontFormatting.IsFontCancellationModified=(value); }
        }

        /**
         * @return true if font outline type was modified from default   
         */
        public bool IsFontOutlineModified
        {
            get
            {
                return fontFormatting.IsFontOutlineModified;
            }
            set { fontFormatting.IsFontOutlineModified=(value); }
        }

        /**
         * @return true if font shadow type was modified from default   
         */
        public bool IsFontShadowModified
        {
            get
            {
                return fontFormatting.IsFontShadowModified;
            }
            set { fontFormatting.IsFontShadowModified=value; }
        }

        /**
         * @return true if font style was modified from default   
         */
        public bool IsFontStyleModified
        {
            get
            {
                return fontFormatting.IsFontStyleModified;
            }
            set { fontFormatting.IsFontStyleModified=value; }
        }

        /**
         * @return true if font style was Set to <i>italic</i> 
         */
        public bool IsItalic
        {
            get { return fontFormatting.IsFontStyleModified && fontFormatting.IsItalic; }
        }

        /**
         * @return true if font outline Is on
         */
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

        /**
         * @return true if font shadow Is on
         */
        public bool IsShadowOn
        {
            get{return fontFormatting.IsFontOutlineModified && fontFormatting.IsShadowOn;}
            set
            {
                fontFormatting.IsShadowOn=value;
                fontFormatting.IsFontShadowModified=value;
            }
        }

        /**
         * @return true if font strikeout Is on
         */
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

        /**
         * @return true if font Underline type was modified from default   
         */
        public bool IsUnderlineTypeModified
        {
            get{return fontFormatting.IsUnderlineTypeModified;}
            set { fontFormatting.IsUnderlineTypeModified=value; }
        }

        /**
         * @return true if font weight was modified from default   
         */
        public bool IsFontWeightModified
        {
            get{
                return fontFormatting.IsFontWeightModified;
            }

        }

        /**
         * Set font style options.
         * 
         * @param italic - if true, Set posture style to italic, otherwise to normal 
         * @param bold- if true, Set font weight to bold, otherwise to normal
         */

        public void SetFontStyle(bool italic, bool bold)
        {
            bool modified = italic || bold;
            fontFormatting.IsItalic=italic;
            fontFormatting.IsBold=bold;
            fontFormatting.IsFontStyleModified=modified;
            fontFormatting.IsFontWeightModified=modified;
        }

        /**
         * Set font style options to default values (non-italic, non-bold)
         */
        public void ResetFontStyle()
        {
            SetFontStyle(false, false);
        }

 



    }
}
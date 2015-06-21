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

namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;
    using NPOI.Util;

    /**
     * @author Philipp Epp
     *
     */
    public class XWPFAbstractNum
    {
        private CT_AbstractNum ctAbstractNum;
        protected XWPFNumbering numbering;

        protected XWPFAbstractNum()
        {
            this.ctAbstractNum = null;
            this.numbering = null;

        }
        public XWPFAbstractNum(CT_AbstractNum abstractNum)
        {
            this.ctAbstractNum = abstractNum;
        }

        public XWPFAbstractNum(CT_AbstractNum ctAbstractNum, XWPFNumbering numbering)
        {
            this.ctAbstractNum = ctAbstractNum;
            this.numbering = numbering;
        }
        public CT_AbstractNum GetAbstractNum()
        {
            return ctAbstractNum;
        }

        public XWPFNumbering GetNumbering()
        {
            return numbering;
        }

        public CT_AbstractNum GetCTAbstractNum()
        {
            return ctAbstractNum;
        }

        public void SetNumbering(XWPFNumbering numbering)
        {
            this.numbering = numbering;
        }
        #region AbstractNum property
        /// <summary>
        /// Abstract Numbering Definition Type
        /// </summary>
        public MultiLevelType MultiLevelType
        {
            get { return EnumConverter.ValueOf<MultiLevelType, ST_MultiLevelType>(ctAbstractNum.multiLevelType.val); }
            set { ctAbstractNum.multiLevelType.val = EnumConverter.ValueOf<ST_MultiLevelType, MultiLevelType>(value); }
        }
        public string AbstractNumId
        {
            get { return ctAbstractNum.abstractNumId; }
            set { ctAbstractNum.abstractNumId = value; }
        }
        //square, dot, diamond
        private char[] lvlText = new char[] { '\u006e', '\u006c', '\u0075' };

        internal void InitLvl()
        {
            List<CT_Lvl> list = new List<CT_Lvl>();
            for (int i = 0; i < 9; i++)
            {
                CT_Lvl lvl = new CT_Lvl();
                lvl.start.val = "1";
                lvl.tentative = i==0? ST_OnOff.on : ST_OnOff.off;
                lvl.ilvl = i.ToString();
                lvl.lvlJc.val = ST_Jc.left;
                lvl.numFmt.val = ST_NumberFormat.bullet;
                lvl.lvlText.val = lvlText[i % 3].ToString();
                CT_Ind ind = lvl.pPr.AddNewInd();
                ind.left = (420 * (i + 1)).ToString();
                ind.hanging = 420;
                CT_Fonts fonts = lvl.rPr.AddNewRFonts();
                fonts.ascii = "Wingdings";
                fonts.hAnsi = "Wingdings";
                fonts.hint = ST_Hint.@default;

                list.Add(lvl);
            }
            ctAbstractNum.lvl = list;
        }
        #endregion

        internal void SetLevelTentative(int lvl, bool tentative)
        {
            if (tentative)
                this.ctAbstractNum.lvl[lvl].tentative = ST_OnOff.on;
            else
                this.ctAbstractNum.lvl[lvl].tentative = ST_OnOff.off;
        }
    }
    /// <summary>
    /// Numbering Definition Type
    /// </summary>
    public enum MultiLevelType
    {
        /// <summary>
        /// Single Level Numbering Definition
        /// </summary>
        SingleLevel,

        /// <summary>
        /// Multilevel Numbering Definition
        /// </summary>
        Multilevel,

        /// <summary>
        /// Hybrid Multilevel Numbering Definition
        /// </summary>
        HybridMultilevel
    }
    /// <summary>
    /// Numbering Format
    /// </summary>
    public enum NumberFormat
    {
        /// <summary>
        /// Decimal Numbers
        /// </summary>
        Decimal,

        /// <summary>
        /// Uppercase Roman Numerals
        /// </summary>
        UpperRoman,

        /// <summary>
        /// Lowercase Roman Numerals
        /// </summary>
        LowerRoman,

        /// <summary>
        /// Uppercase Latin Alphabet
        /// </summary>
        UpperLetter,

        /// <summary>
        /// Lowercase Latin Alphabet
        /// </summary>
        LowerLetter,

        /// <summary>
        /// Ordinal
        /// </summary>
        Ordinal,

        /// <summary>
        /// Cardinal Text
        /// </summary>
        CardinalText,

        /// <summary>
        /// Ordinal Text
        /// </summary>
        OrdinalText,

        /// <summary>
        /// Hexadecimal Numbering
        /// </summary>
        Hex,

        /// <summary>
        /// Chicago Manual of Style
        /// </summary>
        Chicago,

        /// <summary>
        /// Ideographs
        /// </summary>
        IdeographDigital,

        /// <summary>
        /// Japanese Counting System
        /// </summary>
        JapaneseCounting,

        /// <summary>
        /// AIUEO Order Hiragana
        /// </summary>
        Aiueo,

        /// <summary>
        /// Iroha Ordered Katakana
        /// </summary>
        Iroha,

        /// <summary>
        /// Double Byte Arabic Numerals
        /// </summary>
        DecimalFullWidth,

        /// <summary>
        /// Single Byte Arabic Numerals
        /// </summary>
        DecimalHalfWidth,

        /// <summary>
        /// Japanese Legal Numbering
        /// </summary>
        JapaneseLegal,

        /// <summary>
        /// Japanese Digital Ten Thousand Counting System
        /// </summary>
        JapaneseDigitalTenThousand,

        /// <summary>
        /// Decimal Numbers Enclosed in a Circle
        /// </summary>
        DecimalEnclosedCircle,

        /// <summary>
        /// Double Byte Arabic Numerals Alternate
        /// </summary>
        DecimalFullWidth2,

        /// <summary>
        /// Full-Width AIUEO Order Hiragana
        /// </summary>
        AiueoFullWidth,

        /// <summary>
        /// Full-Width Iroha Ordered Katakana
        /// </summary>
        IrohaFullWidth,

        /// <summary>
        /// Initial Zero Arabic Numerals
        /// </summary>
        DecimalZero,

        /// <summary>
        /// Bullet
        /// </summary>
        Bullet,

        /// <summary>
        /// Korean Ganada Numbering
        /// </summary>
        Ganada,

        /// <summary>
        /// Korean Chosung Numbering
        /// </summary>
        Chosung,

        /// <summary>
        /// Decimal Numbers Followed by a Period
        /// </summary>
        DecimalEnclosedFullstop,

        /// <summary>
        /// Decimal Numbers Enclosed in Parenthesis
        /// </summary>
        DecimalEnclosedParen,

        /// <summary>
        /// Decimal Numbers Enclosed in a Circle
        /// </summary>
        DecimalEnclosedCircleChinese,

        /// <summary>
        /// Ideographs Enclosed in a Circle
        /// </summary>
        IdeographEnclosedCircle,

        /// <summary>
        /// Traditional Ideograph Format
        /// </summary>
        IdeographTraditional,

        /// <summary>
        /// Zodiac Ideograph Format
        /// </summary>
        IdeographZodiac,

        /// <summary>
        /// Traditional Zodiac Ideograph Format
        /// </summary>
        IdeographZodiacTraditional,

        /// <summary>
        /// Taiwanese Counting System
        /// </summary>
        TaiwaneseCounting,

        /// <summary>
        /// Traditional Legal Ideograph Format
        /// </summary>
        IdeographLegalTraditional,

        /// <summary>
        /// Taiwanese Counting Thousand System
        /// </summary>
        TaiwaneseCountingThousand,

        /// <summary>
        /// Taiwanese Digital Counting System
        /// </summary>
        TaiwaneseDigital,

        /// <summary>
        /// Chinese Counting System
        /// </summary>
        ChineseCounting,

        /// <summary>
        /// Chinese Legal Simplified Format
        /// </summary>
        ChineseLegalSimplified,

        /// <summary>
        /// Chinese Counting Thousand System
        /// </summary>
        ChineseCountingThousand,

        /// <summary>
        /// Korean Digital Counting System
        /// </summary>
        KoreanDigital,

        /// <summary>
        /// Korean Counting System
        /// </summary>
        KoreanCounting,

        /// <summary>
        /// Korean Legal Numbering
        /// </summary>
        KoreanLegal,

        /// <summary>
        /// Korean Digital Counting System Alternate
        /// </summary>
        KoreanDigital2,

        /// <summary>
        /// Vietnamese Numerals
        /// </summary>
        VietnameseCounting,

        /// <summary>
        /// Lowercase Russian Alphabet
        /// </summary>
        RussianLower,

        /// <summary>
        /// Uppercase Russian Alphabet
        /// </summary>
        RussianUpper,

        /// <summary>
        /// No Numbering
        /// </summary>
        None,

        /// <summary>
        /// Number With Dashes
        /// </summary>
        NumberInDash,

        /// <summary>
        /// Hebrew Numerals
        /// </summary>
        Hebrew1,

        /// <summary>
        /// Hebrew Alphabet
        /// </summary>
        Hebrew2,

        /// <summary>
        /// Arabic Alphabet
        /// </summary>
        ArabicAlpha,

        /// <summary>
        /// Arabic Abjad Numerals
        /// </summary>
        ArabicAbjad,

        /// <summary>
        /// Hindi Vowels
        /// </summary>
        HindiVowels,

        /// <summary>
        /// Hindi Consonants
        /// </summary>
        HindiConsonants,

        /// <summary>
        /// Hindi Numbers
        /// </summary>
        HindiNumbers,

        /// <summary>
        /// Hindi Counting System
        /// </summary>
        HindiCounting,

        /// <summary>
        /// Thai Letters
        /// </summary>
        ThaiLetters,

        /// <summary>
        /// Thai Numerals
        /// </summary>
        ThaiNumbers,

        /// <summary>
        /// Thai Counting System
        /// </summary>
        ThaiCounting,
    }
}
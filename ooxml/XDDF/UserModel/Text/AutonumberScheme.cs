/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Text
{

    using NPOI.OpenXmlFormats.Dml;

    public enum AutonumberScheme
    {
        AlphabeticLowercaseParenthesesBoth,
        AlphabeticLowercaseParenthesisRight,
        AlphabeticLowercasePeriod,
        AlphabeticUppercaseParenthesesBoth,
        AlphabeticUppercaseParenthesisRight,
        AlphabeticUppercasePeriod,
        Arabic1Minus,
        Arabic2Minus,
        ArabicDoubleBytePeriod,
        ArabicDoubleBytePlain,
        ArabicParenthesesBoth,
        ArabicParenthesisRight,
        ArabicPeriod,
        ArabicPlain,
        CircleNumberDoubleBytePlain,
        CircleNumberWingdingsBlackPlain,
        CircleNumberWingdingsWhitePlain,
        EastAsianChineseSimplifiedPeriod,
        EastAsianChineseSimplifiedPlain,
        EastAsianChineseTraditionalPeriod,
        EastAsianChineseTraditionalPlain,
        EastAsianJapaneseDoubleBytePeriod,
        EastAsianJapaneseKoreanPeriod,
        EastAsianJapaneseKoreanPlain,
        Hebrew2Minus,
        HindiAlpha1Period,
        HindiAlphaPeriod,
        HindiNumberParenthesisRight,
        HindiNumberPeriod,
        RomanLowercaseParenthesesBoth,
        RomanLowercaseParenthesisRight,
        RomanLowercasePeriod,
        RomanUppercaseParenthesesBoth,
        RomanUppercaseParenthesisRight,
        RomanUppercasePeriod,
        ThaiAlphabeticParenthesesBoth,
        ThaiAlphabeticParenthesisRight,
        ThaiAlphabeticPeriod,
        ThaiNumberParenthesesBoth,
        ThaiNumberParenthesisRight,
        ThaiNumberPeriod
    }
    public static class AutonumberSchemeExtensions
    {
        private static Dictionary<ST_TextAutonumberScheme, AutonumberScheme> reverse = new Dictionary<ST_TextAutonumberScheme, AutonumberScheme>()
        {
            { ST_TextAutonumberScheme.alphaLcParenBoth, AutonumberScheme.AlphabeticLowercaseParenthesesBoth },
            { ST_TextAutonumberScheme.alphaLcParenR, AutonumberScheme.AlphabeticLowercaseParenthesisRight },
            { ST_TextAutonumberScheme.alphaLcPeriod, AutonumberScheme.AlphabeticLowercasePeriod },
            { ST_TextAutonumberScheme.alphaUcParenBoth, AutonumberScheme.AlphabeticUppercaseParenthesesBoth },
            { ST_TextAutonumberScheme.alphaUcParenR, AutonumberScheme.AlphabeticUppercaseParenthesisRight },
            { ST_TextAutonumberScheme.alphaUcPeriod, AutonumberScheme.AlphabeticUppercasePeriod },
            { ST_TextAutonumberScheme.arabic1Minus, AutonumberScheme.Arabic1Minus },
            { ST_TextAutonumberScheme.arabic2Minus, AutonumberScheme.Arabic2Minus },
            { ST_TextAutonumberScheme.arabicDbPeriod, AutonumberScheme.ArabicDoubleBytePeriod },
            { ST_TextAutonumberScheme.arabicDbPlain, AutonumberScheme.ArabicDoubleBytePlain },
            { ST_TextAutonumberScheme.arabicParenBoth, AutonumberScheme.ArabicParenthesesBoth },
            { ST_TextAutonumberScheme.arabicParenR, AutonumberScheme.ArabicParenthesisRight },
            { ST_TextAutonumberScheme.arabicPeriod, AutonumberScheme.ArabicPeriod },
            { ST_TextAutonumberScheme.arabicPlain, AutonumberScheme.ArabicPlain },
            { ST_TextAutonumberScheme.circleNumDbPlain, AutonumberScheme.CircleNumberDoubleBytePlain },
            { ST_TextAutonumberScheme.circleNumWdBlackPlain, AutonumberScheme.CircleNumberWingdingsBlackPlain },
            { ST_TextAutonumberScheme.circleNumWdWhitePlain, AutonumberScheme.CircleNumberWingdingsWhitePlain },
            { ST_TextAutonumberScheme.ea1ChsPeriod, AutonumberScheme.EastAsianChineseSimplifiedPeriod },
            { ST_TextAutonumberScheme.ea1ChsPlain, AutonumberScheme.EastAsianChineseSimplifiedPlain },
            { ST_TextAutonumberScheme.ea1ChtPeriod, AutonumberScheme.EastAsianChineseTraditionalPeriod },
            { ST_TextAutonumberScheme.ea1ChtPlain, AutonumberScheme.EastAsianChineseTraditionalPlain },
            { ST_TextAutonumberScheme.ea1JpnChsDbPeriod, AutonumberScheme.EastAsianJapaneseDoubleBytePeriod },
            { ST_TextAutonumberScheme.ea1JpnKorPeriod, AutonumberScheme.EastAsianJapaneseKoreanPeriod },
            { ST_TextAutonumberScheme.ea1JpnKorPlain, AutonumberScheme.EastAsianJapaneseKoreanPlain },
            { ST_TextAutonumberScheme.hebrew2Minus, AutonumberScheme.Hebrew2Minus },
            { ST_TextAutonumberScheme.hindiAlpha1Period, AutonumberScheme.HindiAlpha1Period },
            { ST_TextAutonumberScheme.hindiAlphaPeriod, AutonumberScheme.HindiAlphaPeriod },
            { ST_TextAutonumberScheme.hindiNumParenR, AutonumberScheme.HindiNumberParenthesisRight },
            { ST_TextAutonumberScheme.hindiNumPeriod, AutonumberScheme.HindiNumberPeriod },
            { ST_TextAutonumberScheme.romanLcParenBoth, AutonumberScheme.RomanLowercaseParenthesesBoth },
            { ST_TextAutonumberScheme.romanLcParenR, AutonumberScheme.RomanLowercaseParenthesisRight },
            { ST_TextAutonumberScheme.romanLcPeriod, AutonumberScheme.RomanLowercasePeriod },
            { ST_TextAutonumberScheme.romanUcParenBoth, AutonumberScheme.RomanUppercaseParenthesesBoth },
            { ST_TextAutonumberScheme.romanUcParenR, AutonumberScheme.RomanUppercaseParenthesisRight },
            { ST_TextAutonumberScheme.romanUcPeriod, AutonumberScheme.RomanUppercasePeriod },
            { ST_TextAutonumberScheme.thaiAlphaParenBoth, AutonumberScheme.ThaiAlphabeticParenthesesBoth },
            { ST_TextAutonumberScheme.thaiAlphaParenR, AutonumberScheme.ThaiAlphabeticParenthesisRight },
            { ST_TextAutonumberScheme.thaiAlphaPeriod, AutonumberScheme.ThaiAlphabeticPeriod },
            { ST_TextAutonumberScheme.thaiNumParenBoth, AutonumberScheme.ThaiNumberParenthesesBoth },
            { ST_TextAutonumberScheme.thaiNumParenR, AutonumberScheme.ThaiNumberParenthesisRight },
            { ST_TextAutonumberScheme.thaiNumPeriod, AutonumberScheme.ThaiNumberPeriod },
        };
        public static AutonumberScheme ValueOf(ST_TextAutonumberScheme value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextAutonumberScheme", nameof(value));
        }
        public static ST_TextAutonumberScheme ToST_TextAutonumberScheme(this AutonumberScheme value)
        {
            switch(value)
            {
                case AutonumberScheme.AlphabeticLowercaseParenthesesBoth:
                    return ST_TextAutonumberScheme.alphaLcParenBoth;
                case AutonumberScheme.AlphabeticLowercaseParenthesisRight:
                    return ST_TextAutonumberScheme.alphaLcParenR;
                case AutonumberScheme.AlphabeticLowercasePeriod:
                    return ST_TextAutonumberScheme.alphaLcPeriod;
                case AutonumberScheme.AlphabeticUppercaseParenthesesBoth:
                    return ST_TextAutonumberScheme.alphaUcParenBoth;
                case AutonumberScheme.AlphabeticUppercaseParenthesisRight:
                    return ST_TextAutonumberScheme.alphaUcParenR;
                case AutonumberScheme.AlphabeticUppercasePeriod:
                    return ST_TextAutonumberScheme.alphaUcPeriod;
                case AutonumberScheme.Arabic1Minus:
                    return ST_TextAutonumberScheme.arabic1Minus;
                case AutonumberScheme.Arabic2Minus:
                    return ST_TextAutonumberScheme.arabic2Minus;
                case AutonumberScheme.ArabicDoubleBytePeriod:
                    return ST_TextAutonumberScheme.arabicDbPeriod;
                case AutonumberScheme.ArabicDoubleBytePlain:
                    return ST_TextAutonumberScheme.arabicDbPlain;
                case AutonumberScheme.ArabicParenthesesBoth:
                    return ST_TextAutonumberScheme.arabicParenBoth;
                case AutonumberScheme.ArabicParenthesisRight:
                    return ST_TextAutonumberScheme.arabicParenR;
                case AutonumberScheme.ArabicPeriod:
                    return ST_TextAutonumberScheme.arabicPeriod;
                case AutonumberScheme.ArabicPlain:
                    return ST_TextAutonumberScheme.arabicPlain;
                case AutonumberScheme.CircleNumberDoubleBytePlain:
                    return ST_TextAutonumberScheme.circleNumDbPlain;
                case AutonumberScheme.CircleNumberWingdingsBlackPlain:
                    return ST_TextAutonumberScheme.circleNumWdBlackPlain;
                case AutonumberScheme.CircleNumberWingdingsWhitePlain:
                    return ST_TextAutonumberScheme.circleNumWdWhitePlain;
                case AutonumberScheme.EastAsianChineseSimplifiedPeriod:
                    return ST_TextAutonumberScheme.ea1ChsPeriod;
                case AutonumberScheme.EastAsianChineseSimplifiedPlain:
                    return ST_TextAutonumberScheme.ea1ChsPlain;
                case AutonumberScheme.EastAsianChineseTraditionalPeriod:
                    return ST_TextAutonumberScheme.ea1ChtPeriod;
                case AutonumberScheme.EastAsianChineseTraditionalPlain:
                    return ST_TextAutonumberScheme.ea1ChtPlain;
                case AutonumberScheme.EastAsianJapaneseDoubleBytePeriod:
                    return ST_TextAutonumberScheme.ea1JpnChsDbPeriod;
                case AutonumberScheme.EastAsianJapaneseKoreanPeriod:
                    return ST_TextAutonumberScheme.ea1JpnKorPeriod;
                case AutonumberScheme.EastAsianJapaneseKoreanPlain:
                    return ST_TextAutonumberScheme.ea1JpnKorPlain;
                case AutonumberScheme.Hebrew2Minus:
                    return ST_TextAutonumberScheme.hebrew2Minus;
                case AutonumberScheme.HindiAlpha1Period:
                    return ST_TextAutonumberScheme.hindiAlpha1Period;
                case AutonumberScheme.HindiAlphaPeriod:
                    return ST_TextAutonumberScheme.hindiAlphaPeriod;
                case AutonumberScheme.HindiNumberParenthesisRight:
                    return ST_TextAutonumberScheme.hindiNumParenR;
                case AutonumberScheme.HindiNumberPeriod:
                    return ST_TextAutonumberScheme.hindiNumPeriod;
                case AutonumberScheme.RomanLowercaseParenthesesBoth:
                    return ST_TextAutonumberScheme.romanLcParenBoth;
                case AutonumberScheme.RomanLowercaseParenthesisRight:
                    return ST_TextAutonumberScheme.romanLcParenR;
                case AutonumberScheme.RomanLowercasePeriod:
                    return ST_TextAutonumberScheme.romanLcPeriod;
                case AutonumberScheme.RomanUppercaseParenthesesBoth:
                    return ST_TextAutonumberScheme.romanUcParenBoth;
                case AutonumberScheme.RomanUppercaseParenthesisRight:
                    return ST_TextAutonumberScheme.romanUcParenR;
                case AutonumberScheme.RomanUppercasePeriod:
                    return ST_TextAutonumberScheme.romanUcPeriod;
                case AutonumberScheme.ThaiAlphabeticParenthesesBoth:
                    return ST_TextAutonumberScheme.thaiAlphaParenBoth;
                case AutonumberScheme.ThaiAlphabeticParenthesisRight:
                    return ST_TextAutonumberScheme.thaiAlphaParenR;
                case AutonumberScheme.ThaiAlphabeticPeriod:
                    return ST_TextAutonumberScheme.thaiAlphaPeriod;
                case AutonumberScheme.ThaiNumberParenthesesBoth:
                    return ST_TextAutonumberScheme.thaiNumParenBoth;
                case AutonumberScheme.ThaiNumberParenthesisRight:
                    return ST_TextAutonumberScheme.thaiNumParenR;
                case AutonumberScheme.ThaiNumberPeriod:
                    return ST_TextAutonumberScheme.thaiNumPeriod;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



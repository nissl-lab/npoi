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
using System.Runtime.InteropServices;
using System.Text;

namespace NPOI.Common.UserModel.Fonts
{
    using NPOI.Util;
    using System.Runtime.InteropServices;



    /// <summary>
    /// Charset represents the basic Set of characters associated with a font (that it can display), and
    /// corresponds to the ANSI codepage (8-bit or DBCS) of that character Set used by a given language.
    /// </summary>
    /// <remarks>
    /// @since POI 3.17-beta2
    /// </remarks>
    public enum FontCharset
    {
        /** Specifies the English character Set. */
        ANSI=0x00000000,
        /**
         * Specifies a character Set based on the current system locale;
         * for example, when the system locale is United States English,
         * the default character Set is ANSI_CHARSET.
         */
        DEFAULT=0x00000001,
        /** Specifies a character Set of symbols. */
        SYMBOL=0x00000002,
        /** Specifies the Apple Macintosh character Set. */
        MAC=0x0000004D,
        /** Specifies the Japanese character Set. */
        SHIFTJIS=0x00000080,
        /** Also spelled "Hangeul". Specifies the Hangul Korean character Set. */
        HANGUL=0x00000081,
        /** Also spelled "Johap". Specifies the Johab Korean character Set. */
        JOHAB=0x00000082,
        /** Specifies the "simplified" Chinese character Set for People's Republic of China. */
        GB2312=0x00000086,
        /**
         * Specifies the "traditional" Chinese character Set, used mostly in
         * Taiwan and in the Hong Kong and Macao Special Administrative Regions.
         */
        CHINESEBIG5=0x00000088,
        /** Specifies the Greek character Set. */
        GREEK=0x000000A1,
        /** Specifies the Turkish character Set. */
        TURKISH=0x000000A2,
        /** Specifies the Vietnamese character Set. */
        VIETNAMESE=0x000000A3,
        /** Specifies the Hebrew character Set. */
        HEBREW=0x000000B1,
        /** Specifies the Arabic character Set. */
        ARABIC=0x000000B2,
        /** Specifies the Baltic (Northeastern European) character Set. */
        BALTIC=0x000000BA,
        /** Specifies the Russian Cyrillic character Set. */
        RUSSIAN=0x000000CC,
        /** Specifies the Thai character Set. */
        THAI_=0x000000DE,
        /** Specifies a Eastern European character Set. */
        EASTEUROPE=0x000000EE,
        /**
         * Specifies a mapping to one of the OEM code pages,
         * according to the current system locale Setting.
         */
        OEM=0x000000FF,
    }
    public class FontCharsetInfo
    {
        private static FontCharset[] _table = new FontCharset[256];
        /** Specifies the English character Set. */
        public static FontCharsetInfo ANSI = new FontCharsetInfo(0x00000000, "Cp1252");
        /**
         * Specifies a character Set based on the current system locale;
         * for example, when the system locale is United States English,
         * the default character Set is ANSI_CHARSET.
         */
        public static FontCharsetInfo DEFAULT = new FontCharsetInfo(0x00000001, "Cp1252");
        /** Specifies a character Set of symbols. */
        public static FontCharsetInfo SYMBOL = new FontCharsetInfo(0x00000002, "");
        /** Specifies the Apple Macintosh character Set. */
        public static FontCharsetInfo MAC = new FontCharsetInfo(0x0000004D, "MacRoman");
        /** Specifies the Japanese character Set. */
        public static FontCharsetInfo SHIFTJIS = new FontCharsetInfo(0x00000080, "Shift_JIS");
        /** Also spelled "Hangeul". Specifies the Hangul Korean character Set. */
        public static FontCharsetInfo HANGUL = new FontCharsetInfo(0x00000081, "cp949");
        /** Also spelled "Johap". Specifies the Johab Korean character Set. */
        public static FontCharsetInfo JOHAB = new FontCharsetInfo(0x00000082, "x-Johab");
        /** Specifies the "simplified" Chinese character Set for People's Republic of China. */
        public static FontCharsetInfo GB2312 = new FontCharsetInfo(0x00000086, "GB2312");
        /**
         * Specifies the "traditional" Chinese character Set, used mostly in
         * Taiwan and in the Hong Kong and Macao Special Administrative Regions.
         */
        public static FontCharsetInfo CHINESEBIG5 = new FontCharsetInfo(0x00000088, "Big5");
        /** Specifies the Greek character Set. */
        public static FontCharsetInfo GREEK = new FontCharsetInfo(0x000000A1, "Cp1253");
        /** Specifies the Turkish character Set. */
        public static FontCharsetInfo TURKISH = new FontCharsetInfo(0x000000A2, "Cp1254");
        /** Specifies the Vietnamese character Set. */
        public static FontCharsetInfo VIETNAMESE = new FontCharsetInfo(0x000000A3, "Cp1258");
        /** Specifies the Hebrew character Set. */
        public static FontCharsetInfo HEBREW = new FontCharsetInfo(0x000000B1, "Cp1255");
        /** Specifies the Arabic character Set. */
        public static FontCharsetInfo ARABIC = new FontCharsetInfo(0x000000B2, "Cp1256");
        /** Specifies the Baltic (Northeastern European) character Set. */
        public static FontCharsetInfo BALTIC = new FontCharsetInfo(0x000000BA, "Cp1257");
        /** Specifies the Russian Cyrillic character Set. */
        public static FontCharsetInfo RUSSIAN = new FontCharsetInfo(0x000000CC, "Cp1251");
        /** Specifies the Thai character Set. */
        public static FontCharsetInfo THAI_ = new FontCharsetInfo(0x000000DE, "x-windows-874");
        /** Specifies a Eastern European character Set. */
        public static FontCharsetInfo EASTEUROPE = new FontCharsetInfo(0x000000EE, "Cp1250");
        /**
         * Specifies a mapping to one of the OEM code pages,
         * according to the current system locale Setting.
         */
        public static FontCharsetInfo OEM = new FontCharsetInfo(0x000000FF, "Cp1252");
        static FontCharsetInfo()
        {
            _table[0x00000000] = FontCharset.ANSI;
            _table[0x00000001] = FontCharset.DEFAULT;
            _table[0x00000002] = FontCharset.SYMBOL;
            _table[0x0000004D] = FontCharset.MAC;
            _table[0x00000080] = FontCharset.SHIFTJIS;
            _table[0x00000081] = FontCharset.HANGUL;
            _table[0x00000082] = FontCharset.JOHAB;
            _table[0x00000086] = FontCharset.GB2312;
            _table[0x00000088] = FontCharset.CHINESEBIG5;
            _table[0x000000A1] = FontCharset.GREEK;
            _table[0x000000A2] = FontCharset.TURKISH;
            _table[0x000000A3] = FontCharset.VIETNAMESE;
            _table[0x000000B1] = FontCharset.HEBREW;
            _table[0x000000B2] = FontCharset.ARABIC;
            _table[0x000000BA] = FontCharset.BALTIC;
            _table[0x000000CC] = FontCharset.RUSSIAN;
            _table[0x000000DE] = FontCharset.THAI_;
            _table[0x000000EE] = FontCharset.EASTEUROPE;
            _table[0x000000FF] = FontCharset.OEM;

        }
        public int NativeId { get; set; }
        
        public Encoding Charset { get; set; }
        public FontCharsetInfo(int flag, string charsetName)
        {
            this.NativeId = flag;
            if(charsetName.Length > 0)
            {
                try
                {
                    Charset = Encoding.GetEncoding(charsetName);
                    return;
                }
                catch(Exception e)
                {

                }
            }
            Charset = null;
        }

        public static FontCharset? ValueOf(int value)
        {
            return (value < 0 || value >= _table.Length) ? null : _table[value];
        }
    }
}

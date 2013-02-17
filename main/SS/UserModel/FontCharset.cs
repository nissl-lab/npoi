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

namespace NPOI.SS.UserModel
{


    /**
     * Charset represents the basic set of characters associated with a font (that it can display), and 
     * corresponds to the ANSI codepage (8-bit or DBCS) of that character set used by a given language. 
     * 
     * @author Gisella Bronzetti
     */
    public class FontCharset
    {
        public static readonly FontCharset ANSI = new FontCharset(0);
        public static readonly FontCharset DEFAULT = new FontCharset(1);
        public static readonly FontCharset SYMBOL = new FontCharset(2);
        public static readonly FontCharset MAC = new FontCharset(77);
        public static readonly FontCharset SHIFTJIS = new FontCharset(128);
        public static readonly FontCharset HANGEUL = new FontCharset(129);
        public static readonly FontCharset JOHAB = new FontCharset(130);
        public static readonly FontCharset GB2312 = new FontCharset(134);
        public static readonly FontCharset CHINESEBIG5 = new FontCharset(136);
        public static readonly FontCharset GREEK = new FontCharset(161);
        public static readonly FontCharset TURKISH = new FontCharset(162);
        public static readonly FontCharset VIETNAMESE = new FontCharset(163);
        public static readonly FontCharset HEBREW = new FontCharset(177);
        public static readonly FontCharset ARABIC = new FontCharset(178);
        public static readonly FontCharset BALTIC = new FontCharset(186);
        public static readonly FontCharset RUSSIAN = new FontCharset(204);
        public static readonly FontCharset THAI = new FontCharset(222);
        public static readonly FontCharset EASTEUROPE = new FontCharset(238);
        public static readonly FontCharset OEM = new FontCharset(255);


        private int charset;

        private FontCharset(int value)
        {
            charset = value;
        }

        /**
         * Returns value of this charset
         *
         * @return value of this charset
         */
        public int Value
        {
            get
            {
                return charset;
            }
        }

        private static FontCharset[] _table = null;

        static FontCharset()
        {
            if (_table == null)
            {
                _table = new FontCharset[256];
                _table[0] = FontCharset.ANSI;
                _table[1] = FontCharset.DEFAULT;
                _table[2] = FontCharset.SYMBOL;
                _table[77] = FontCharset.MAC;
                _table[128] = FontCharset.SHIFTJIS;
                _table[129] = FontCharset.HANGEUL;
                _table[130] = FontCharset.JOHAB;
                _table[134] = FontCharset.GB2312;
                _table[136] = FontCharset.CHINESEBIG5;
                _table[161] = FontCharset.GREEK;
                _table[162] = FontCharset.TURKISH;
                _table[163] = FontCharset.VIETNAMESE;
                _table[177] = FontCharset.HEBREW;
                _table[178] = FontCharset.ARABIC;
                _table[186] = FontCharset.BALTIC;
                _table[204] = FontCharset.RUSSIAN;
                _table[222] = FontCharset.THAI;
                _table[238] = FontCharset.EASTEUROPE;
                _table[255] = FontCharset.OEM;
            }
        }

        public static FontCharset ValueOf(int value)
        {
            if(value>=0&&value<=255)
                return _table[value];
            return null;
        }
    }
}
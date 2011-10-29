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
public class FontCharset {

    public static FontCharset ANSI = new FontCharset(0);
    public static FontCharset DEFAULT = new FontCharset(1);
    public static FontCharset SYMBOL = new FontCharset(2);
    public static FontCharset MAC = new FontCharset(77);
    public static FontCharset SHIFTJIS = new FontCharset(128);
    public static FontCharset HANGEUL = new FontCharset(129);
    public static FontCharset JOHAB = new FontCharset(130);
    public static FontCharset GB2312 = new FontCharset(134);
    public static FontCharset CHINESEBIG5 = new FontCharset(136);
    public static FontCharset GREEK = new FontCharset(161);
    public static FontCharset TURKISH = new FontCharset(162);
    public static FontCharset VIETNAMESE = new FontCharset(163);
    public static FontCharset HEBREW = new FontCharset(177);
    public static FontCharset ARABIC = new FontCharset(178);
    public static FontCharset BALTIC = new FontCharset(186);
    public static FontCharset RUSSIAN = new FontCharset(204);
    public static FontCharset THAI = new FontCharset(222);
    public static FontCharset EASTEUROPE = new FontCharset(238);
    public static FontCharset OEM = new FontCharset(255);

    
    private int charset;

    private FontCharset(int value){
        charset = value;
        _table[charset] = this;
    }

    /**
     * Returns value of this charset
     *
     * @return value of this charset
     */
    public int Value
    {
        get{
            return charset;
        }
    }

    private static FontCharset[] _table = new FontCharset[256];

    public static FontCharset ValueOf(int value){
        return _table[value];
    }
}
}
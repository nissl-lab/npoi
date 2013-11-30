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

using System;
using System.IO;
using System.Text;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class CodePageString
    {
        //private final static POILogger logger = POILogFactory
        //   .getLogger( CodePageString.class );

        private static String codepageToEncoding(int codepage)
        {
            if (codepage <= 0)
                throw new UnsupportedEncodingException(
                        "Codepage number may not be " + codepage);
            switch ((Constants)codepage)
            {
                case Constants.CP_UTF16:
                    return "UTF-16";
                case Constants.CP_UTF16_BE:
                    return "UTF-16BE";
                case Constants.CP_UTF8:
                    return "UTF-8";
                case Constants.CP_037:
                    return "cp037";
                case Constants.CP_GBK:
                    return "GBK";
                case Constants.CP_MS949:
                    return "ms949";
                case Constants.CP_WINDOWS_1250:
                    return "windows-1250";
                case Constants.CP_WINDOWS_1251:
                    return "windows-1251";
                case Constants.CP_WINDOWS_1252:
                    return "windows-1252";
                case Constants.CP_WINDOWS_1253:
                    return "windows-1253";
                case Constants.CP_WINDOWS_1254:
                    return "windows-1254";
                case Constants.CP_WINDOWS_1255:
                    return "windows-1255";
                case Constants.CP_WINDOWS_1256:
                    return "windows-1256";
                case Constants.CP_WINDOWS_1257:
                    return "windows-1257";
                case Constants.CP_WINDOWS_1258:
                    return "windows-1258";
                case Constants.CP_JOHAB:
                    return "johab";
                case Constants.CP_MAC_ROMAN:
                    return "MacRoman";
                case Constants.CP_MAC_JAPAN:
                    return "SJIS";
                case Constants.CP_MAC_CHINESE_TRADITIONAL:
                    return "Big5";
                case Constants.CP_MAC_KOREAN:
                    return "EUC-KR";
                case Constants.CP_MAC_ARABIC:
                    return "MacArabic";
                case Constants.CP_MAC_HEBREW:
                    return "MacHebrew";
                case Constants.CP_MAC_GREEK:
                    return "MacGreek";
                case Constants.CP_MAC_CYRILLIC:
                    return "MacCyrillic";
                case Constants.CP_MAC_CHINESE_SIMPLE:
                    return "EUC_CN";
                case Constants.CP_MAC_ROMANIA:
                    return "MacRomania";
                case Constants.CP_MAC_UKRAINE:
                    return "MacUkraine";
                case Constants.CP_MAC_THAI:
                    return "MacThai";
                case Constants.CP_MAC_CENTRAL_EUROPE:
                    return "MacCentralEurope";
                case Constants.CP_MAC_ICELAND:
                    return "MacIceland";
                case Constants.CP_MAC_TURKISH:
                    return "MacTurkish";
                case Constants.CP_MAC_CROATIAN:
                    return "MacCroatian";
                case Constants.CP_US_ACSII:
                case Constants.CP_US_ASCII2:
                    return "US-ASCII";
                case Constants.CP_KOI8_R:
                    return "KOI8-R";
                case Constants.CP_ISO_8859_1:
                    return "ISO-8859-1";
                case Constants.CP_ISO_8859_2:
                    return "ISO-8859-2";
                case Constants.CP_ISO_8859_3:
                    return "ISO-8859-3";
                case Constants.CP_ISO_8859_4:
                    return "ISO-8859-4";
                case Constants.CP_ISO_8859_5:
                    return "ISO-8859-5";
                case Constants.CP_ISO_8859_6:
                    return "ISO-8859-6";
                case Constants.CP_ISO_8859_7:
                    return "ISO-8859-7";
                case Constants.CP_ISO_8859_8:
                    return "ISO-8859-8";
                case Constants.CP_ISO_8859_9:
                    return "ISO-8859-9";
                case Constants.CP_ISO_2022_JP1:
                case Constants.CP_ISO_2022_JP2:
                case Constants.CP_ISO_2022_JP3:
                    return "ISO-2022-JP";
                case Constants.CP_ISO_2022_KR:
                    return "ISO-2022-KR";
                case Constants.CP_EUC_JP:
                    return "EUC-JP";
                case Constants.CP_EUC_KR:
                    return "EUC-KR";
                case Constants.CP_GB2312:
                    return "GB2312";
                case Constants.CP_GB18030:
                    return "GB18030";
                case Constants.CP_SJIS:
                    return "SJIS";
                default:
                    return "cp" + codepage;
            }
        }

        private byte[] _value;

        public CodePageString(byte[] data, int startOffset)
        {
            int offset = startOffset;

            int size = LittleEndian.GetInt(data, offset);
            offset += LittleEndian.INT_SIZE;

            _value = LittleEndian.GetByteArray(data, offset, size);
            if (_value[size - 1] != 0)
            {
                // TODO Some files, such as TestVisioWithCodepage.vsd, are currently
                //  triggering this for values that don't look like codepages
                // See Bug #52258 for details
                //logger.log(POILogger.WARN, "CodePageString started at offset #" + offset
                //            + " is not NULL-terminated" );
                //            throw new IllegalPropertySetDataException(
                //                    "CodePageString started at offset #" + offset
                //                            + " is not NULL-terminated" );
            }
        }

        public CodePageString(String aString, int codepage)
        {
            SetJavaValue(aString, codepage);
        }

        public String GetJavaValue(int codepage)
        {
            String result;
            if (codepage == -1)
                result = Encoding.UTF8.GetString(_value);
            else
                result = Encoding.GetEncoding(codepage).GetString(_value);
            int terminator = result.IndexOf('\0');
            if (terminator == -1)
            {
                //logger.log(
                //        POILogger.WARN,
                //        "String terminator (\\0) for CodePageString property value not found."
                //                + "Continue without trimming and hope for the best." );
                return result;
            }
            if (terminator != result.Length - 1)
            {
                //logger.log(
                //        POILogger.WARN,
                //        "String terminator (\\0) for CodePageString property value occured before the end of string. "
                //                + "Trimming and hope for the best." );
            }
            return result.Substring(0, terminator);
        }

        public int Size
        {
            get { return LittleEndian.INT_SIZE + _value.Length; }
        }

        public void SetJavaValue(String aString, int codepage)
        {
            if (codepage == -1)
                _value = Encoding.UTF8.GetBytes(aString + "\0");
            else
                _value = Encoding.GetEncoding(codepage).GetBytes(aString + "\0");
        }

        public int Write(Stream out1)
        {
            LittleEndian.PutInt(_value.Length, out1);
            out1.Write(_value, 0, _value.Length);
            return LittleEndian.INT_SIZE + _value.Length;
        }
    }
}
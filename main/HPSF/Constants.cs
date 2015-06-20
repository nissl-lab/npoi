/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"), you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/
using System;
namespace NPOI.HPSF
{
    /// <summary>
    /// Defines constants of general use.
    /// @author Rainer Klute klute@rainer-klute.de
    /// @since 2004-06-20
    /// </summary>
    [Obsolete("use CodePageUtil")]
    public enum Constants
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        None = 0,
        /** Codepage 037, a special case */
         CP_037 = 37,

        /** Codepage for SJIS */
         CP_SJIS = 932,

        /** Codepage for GBK, aka MS936 */
         CP_GBK = 936,

        /** Codepage for MS949 */
         CP_MS949 = 949,

        /** Codepage for UTF-16 */
         CP_UTF16 = 1200,

        /** Codepage for UTF-16 big-endian */
         CP_UTF16_BE = 1201,

        /** Codepage for Windows 1250 */
         CP_WINDOWS_1250 = 1250,

        /** Codepage for Windows 1251 */
         CP_WINDOWS_1251 = 1251,

        /** Codepage for Windows 1252 */
         CP_WINDOWS_1252 = 1252,

        /** Codepage for Windows 1253 */
         CP_WINDOWS_1253 = 1253,

        /** Codepage for Windows 1254 */
         CP_WINDOWS_1254 = 1254,

        /** Codepage for Windows 1255 */
         CP_WINDOWS_1255 = 1255,

        /** Codepage for Windows 1256 */
         CP_WINDOWS_1256 = 1256,

        /** Codepage for Windows 1257 */
         CP_WINDOWS_1257 = 1257,

        /** Codepage for Windows 1258 */
         CP_WINDOWS_1258 = 1258,

        /** Codepage for Johab */
         CP_JOHAB = 1361,

        /** Codepage for Macintosh Roman (Java: MacRoman) */
         CP_MAC_ROMAN = 10000,

        /** Codepage for Macintosh Japan (Java: unknown - use SJIS, cp942 or
         * cp943) */
         CP_MAC_JAPAN = 10001,

        /** Codepage for Macintosh Chinese Traditional (Java: unknown - use Big5,
         * MS950, or cp937) */
         CP_MAC_CHINESE_TRADITIONAL = 10002,

        /** Codepage for Macintosh Korean (Java: unknown - use EUC_KR or
         * cp949) */
         CP_MAC_KOREAN = 10003,

        /** Codepage for Macintosh Arabic (Java: MacArabic) */
         CP_MAC_ARABIC = 10004,

        /** Codepage for Macintosh Hebrew (Java: MacHebrew) */
         CP_MAC_HEBREW = 10005,

        /** Codepage for Macintosh Greek (Java: MacGreek) */
         CP_MAC_GREEK = 10006,

        /** Codepage for Macintosh Cyrillic (Java: MacCyrillic) */
         CP_MAC_CYRILLIC = 10007,

        /** Codepage for Macintosh Chinese Simplified (Java: unknown - use
         * EUC_CN, ISO2022_CN_GB, MS936 or cp935) */
         CP_MAC_CHINESE_SIMPLE = 10008,

        /** Codepage for Macintosh Romanian (Java: MacRomania) */
         CP_MAC_ROMANIA = 10010,

        /** Codepage for Macintosh Ukrainian (Java: MacUkraine) */
         CP_MAC_UKRAINE = 10017,

        /** Codepage for Macintosh Thai (Java: MacThai) */
         CP_MAC_THAI = 10021,

        /** Codepage for Macintosh Central Europe (Latin-2)
         * (Java: MacCentralEurope) */
         CP_MAC_CENTRAL_EUROPE = 10029,

        /** Codepage for Macintosh Iceland (Java: MacIceland) */
         CP_MAC_ICELAND = 10079,

        /** Codepage for Macintosh Turkish (Java: MacTurkish) */
         CP_MAC_TURKISH = 10081,

        /** Codepage for Macintosh Croatian (Java: MacCroatian) */
         CP_MAC_CROATIAN = 10082,

        /** Codepage for US-ASCII */
         CP_US_ACSII = 20127,

        /** Codepage for KOI8-R */
         CP_KOI8_R = 20866,

        /** Codepage for ISO-8859-1 */
         CP_ISO_8859_1 = 28591,

        /** Codepage for ISO-8859-2 */
         CP_ISO_8859_2 = 28592,

        /** Codepage for ISO-8859-3 */
         CP_ISO_8859_3 = 28593,

        /** Codepage for ISO-8859-4 */
         CP_ISO_8859_4 = 28594,

        /** Codepage for ISO-8859-5 */
         CP_ISO_8859_5 = 28595,

        /** Codepage for ISO-8859-6 */
         CP_ISO_8859_6 = 28596,

        /** Codepage for ISO-8859-7 */
         CP_ISO_8859_7 = 28597,

        /** Codepage for ISO-8859-8 */
         CP_ISO_8859_8 = 28598,

        /** Codepage for ISO-8859-9 */
         CP_ISO_8859_9 = 28599,

        /** Codepage for ISO-2022-JP */
         CP_ISO_2022_JP1 = 50220,

        /** Another codepage for ISO-2022-JP */
         CP_ISO_2022_JP2 = 50221,

        /** Yet another codepage for ISO-2022-JP */
         CP_ISO_2022_JP3 = 50222,

        /** Codepage for ISO-2022-KR */
         CP_ISO_2022_KR = 50225,

        /** Codepage for EUC-JP */
         CP_EUC_JP = 51932,

        /** Codepage for EUC-KR */
         CP_EUC_KR = 51949,

        /** Codepage for GB2312 */
         CP_GB2312 = 52936,

        /** Codepage for GB18030 */
         CP_GB18030 = 54936,

        /** Another codepage for US-ASCII */
         CP_US_ASCII2 = 65000,

        /** Codepage for UTF-8 */
         CP_UTF8 = 65001,

        /** Codepage for Unicode */
         CP_UNICODE = CP_UTF16,
    }
}
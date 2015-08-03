
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
using System.Text;
using System.Collections.Generic;



namespace NPOI.Util
{
    /// <summary>
    /// Title: String Utility Description: Collection of string handling utilities
    /// @author     Andrew C. Oliver
    /// @author     Sergei Kozello (sergeikozello at mail.ru)
    /// @author     Toshiaki Kamoshida (kamoshida.toshiaki at future dot co dot jp)
    /// @since      May 10, 2002
    /// @version    1.0
    /// </summary>
    public class StringUtil
    {
        private static Encoding ISO_8859_1 = Encoding.GetEncoding("ISO-8859-1");
        private static Encoding UTF16LE = Encoding.Unicode;
        private static Dictionary<int, int> msCodepointToUnicode;
        /**     
         *  Constructor for the StringUtil object     
         */
        private StringUtil()
        {
        }

        /// <summary>
        /// Given a byte array of 16-bit unicode characters in Little Endian
        /// Format (most important byte last), return a Java String representation
        /// of it.
        /// { 0x16, 0x00 } -0x16
        /// </summary>
        /// <param name="str">the byte array to be converted</param>
        /// <param name="offset">the initial offset into the
        /// byte array. it is assumed that string[ offset ] and string[ offset + 1 ] contain the first 16-bit unicode character</param>
        /// <param name="len">the Length of the string</param>
        /// <returns>the converted string</returns>                              
        public static String GetFromUnicodeLE(
            byte[] str,
            int offset,
            int len)
        {
            if ((offset < 0) || (offset >= str.Length))
            {
                throw new IndexOutOfRangeException("Illegal offset");
            }
            if ((len < 0) || (((str.Length - offset) / 2) < len))
            {
                throw new ArgumentException("Illegal Length");
            }

            return UTF16LE.GetString(str, offset, len * 2);
        }

        /// <summary>
        /// Given a byte array of 16-bit unicode characters in little endian
        /// Format (most important byte last), return a Java String representation
        /// of it.
        ///{ 0x16, 0x00 } -0x16
        /// </summary>
        /// <param name="str">the byte array to be converted</param>
        /// <returns>the converted string</returns>  
        public static String GetFromUnicodeLE(byte[] str)
        {
            if (str.Length == 0) { return ""; }
            return GetFromUnicodeLE(str, 0, str.Length / 2);
        }

        /**
         * Convert String to 16-bit unicode characters in little endian format
         *
         * @param string the string
         * @return the byte array of 16-bit unicode characters
         */
        public static byte[] GetToUnicodeLE(String string1)
        {
            return UTF16LE.GetBytes(string1);
        }

        /// <summary>
        /// Given a byte array of 16-bit unicode characters in big endian
        /// Format (most important byte first), return a Java String representation
        /// of it.
        ///  { 0x00, 0x16 } -0x16
        /// </summary>
        /// <param name="str">the byte array to be converted</param>
        /// <param name="offset">the initial offset into the
        /// byte array. it is assumed that string[ offset ] and string[ offset + 1 ] contain the first 16-bit unicode character</param>
        /// <param name="len">the Length of the string</param>
        /// <returns> the converted string</returns>
        public static String GetFromUnicodeBE(
            byte[] str,
            int offset,
            int len)
        {
            if ((offset < 0) || (offset >= str.Length))
            {
                throw new IndexOutOfRangeException("Illegal offset");
            }
            if ((len < 0) || (((str.Length - offset) / 2) < len))
            {
                throw new ArgumentException("Illegal Length");
            }
            try
            {
                return Encoding.GetEncoding("UTF-16BE").GetString(str, offset, len * 2);
            }
            catch
            {
                throw new InvalidOperationException(); /*unreachable*/
            }
        }

        /// <summary>
        /// Given a byte array of 16-bit unicode characters in big endian
        /// Format (most important byte first), return a Java String representation
        /// of it.
        /// { 0x00, 0x16 } -0x16
        /// </summary>
        /// <param name="str">the byte array to be converted</param>
        /// <returns>the converted string</returns>      
        public static String GetFromUnicodeBE(byte[] str)
        {
            if (str.Length == 0) { return ""; }
            return GetFromUnicodeBE(str, 0, str.Length / 2);
        }

        /// <summary>
        /// Read 8 bit data (in IsO-8859-1 codepage) into a (unicode) Java
        /// String and return.
        /// (In Excel terms, read compressed 8 bit unicode as a string)
        /// </summary>
        /// <param name="str">byte array to read</param>
        /// <param name="offset">offset to read byte array</param>
        /// <param name="len">Length to read byte array</param>
        /// <returns>generated String instance by reading byte array</returns>
        public static String GetFromCompressedUnicode(
            byte[] str,
            int offset,
            int len)
        {
            int len_to_use = Math.Min(len, str.Length - offset);
            return ISO_8859_1.GetString(str, offset, len_to_use);
        }


        /// <summary>
        /// Takes a unicode (java) string, and returns it as 8 bit data (in IsO-8859-1
        /// codepage).
        /// (In Excel terms, write compressed 8 bit unicode)
        /// </summary>
        /// <param name="input">the String containing the data to be written</param>
        /// <param name="output">the byte array to which the data Is to be written</param>
        /// <param name="offset">an offset into the byte arrat at which the data Is start when written</param>
        public static void PutCompressedUnicode(String input, byte[] output, int offset)
        {
            byte[] bytes = ISO_8859_1.GetBytes(input);
            Array.Copy(bytes, 0, output, offset, bytes.Length);
        }

        public static void PutCompressedUnicode(String input, ILittleEndianOutput out1)
        {
            byte[] bytes = ISO_8859_1.GetBytes(input);
            out1.Write(bytes);
        }


        /// <summary>
        /// Takes a unicode string, and returns it as little endian (most
        /// important byte last) bytes in the supplied byte array.
        /// (In Excel terms, write uncompressed unicode)
        /// </summary>
        /// <param name="input">the String containing the unicode data to be written</param>
        /// <param name="output">the byte array to hold the uncompressed unicode, should be twice the Length of the String</param>
        /// <param name="offset">the offset to start writing into the byte array</param>
        public static void PutUnicodeLE(String input, byte[] output, int offset)
        {
            byte[] bytes = UTF16LE.GetBytes(input);
            Array.Copy(bytes, 0, output, offset, bytes.Length);
        }
        public static void PutUnicodeLE(String input, ILittleEndianOutput out1)
        {
            byte[] bytes = UTF16LE.GetBytes(input);
            out1.Write(bytes);
        }



        /// <summary>
        /// Takes a unicode string, and returns it as big endian (most
        /// important byte first) bytes in the supplied byte array.
        /// (In Excel terms, write uncompressed unicode)
        /// </summary>
        /// <param name="input">the String containing the unicode data to be written</param>
        /// <param name="output">the byte array to hold the uncompressed unicode, should be twice the Length of the String.</param>
        /// <param name="offset">the offset to start writing into the byte array</param>
        public static void PutUnicodeBE(
            String input,
            byte[] output,
            int offset)
        {
            try
            {
                byte[] bytes = Encoding.GetEncoding("UTF-16BE").GetBytes(input);
                Array.Copy(bytes, 0, output, offset, bytes.Length);
            }
            catch
            {
                throw new InvalidOperationException(); /*unreachable*/
            }
        }


        /// <summary>
        /// Gets the preferred encoding.
        /// </summary>
        /// <returns>the encoding we want to use, currently hardcoded to IsO-8859-1</returns>
        public static String GetPreferredEncoding()
        {
            return ISO_8859_1.WebName;
        }

        /// <summary>
        /// check the parameter Has multibyte character
        /// </summary>
        /// <param name="value"> string to check</param>
        /// <returns>
        /// 	<c>true</c> if Has at least one multibyte character; otherwise, <c>false</c>.
        /// </returns>
        public static bool HasMultibyte(String value)
        {
            if (value == null) return false;
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c > 0xFF) return true;
            }
            return false;
        }

        public static String ReadCompressedUnicode(ILittleEndianInput in1, int nChars)
        {
            byte[] buf = new byte[nChars];
            in1.ReadFully(buf);
            return ISO_8859_1.GetString(buf);
        }
        public static String ReadUnicodeLE(ILittleEndianInput in1, int nChars)
        {
            byte[] bytes = new byte[nChars * 2];
            in1.ReadFully(bytes);
            return UTF16LE.GetString(bytes);
        }

        /**
         * InputStream <c>in</c> is expected to contain:
         * <ol>
         * <li>ushort nChars</li>
         * <li>byte is16BitFlag</li>
         * <li>byte[]/char[] characterData</li>
         * </ol>
         * For this encoding, the is16BitFlag is always present even if nChars==0.
         */
        public static String ReadUnicodeString(ILittleEndianInput in1)
        {

            int nChars = in1.ReadUShort();
            byte flag = (byte)in1.ReadByte();
            if ((flag & 0x01) == 0)
            {
                return ReadCompressedUnicode(in1, nChars);
            }
            return ReadUnicodeLE(in1, nChars);
        }
        /**
         * InputStream <c>in</c> is expected to contain:
         * <ol>
         * <li>byte is16BitFlag</li>
         * <li>byte[]/char[] characterData</li>
         * </ol>
         * For this encoding, the is16BitFlag is always present even if nChars==0.
         * <br/>
         * This method should be used when the nChars field is <em>not</em> stored 
         * as a ushort immediately before the is16BitFlag. Otherwise, {@link 
         * #readUnicodeString(LittleEndianInput)} can be used. 
         */
        public static String ReadUnicodeString(ILittleEndianInput in1, int nChars)
        {
            byte is16Bit = (byte)in1.ReadByte();
            if ((is16Bit & 0x01) == 0)
            {
                return ReadCompressedUnicode(in1, nChars);
            }
            return ReadUnicodeLE(in1, nChars);
        }
        /**
         * OutputStream <c>out</c> will get:
         * <ol>
         * <li>ushort nChars</li>
         * <li>byte is16BitFlag</li>
         * <li>byte[]/char[] characterData</li>
         * </ol>
         * For this encoding, the is16BitFlag is always present even if nChars==0.
         */
        public static void WriteUnicodeString(ILittleEndianOutput out1, String value)
        {

            int nChars = value.Length;
            out1.WriteShort(nChars);
            bool is16Bit = HasMultibyte(value);
            out1.WriteByte(is16Bit ? 0x01 : 0x00);
            if (is16Bit)
            {
                PutUnicodeLE(value, out1);
            }
            else
            {
                PutCompressedUnicode(value, out1);
            }
        }
        /**
         * OutputStream <c>out</c> will get:
         * <ol>
         * <li>byte is16BitFlag</li>
         * <li>byte[]/char[] characterData</li>
         * </ol>
         * For this encoding, the is16BitFlag is always present even if nChars==0.
         * <br/>
         * This method should be used when the nChars field is <em>not</em> stored 
         * as a ushort immediately before the is16BitFlag. Otherwise, {@link 
         * #writeUnicodeString(LittleEndianOutput, String)} can be used. 
         */
        public static void WriteUnicodeStringFlagAndData(ILittleEndianOutput out1, String value)
        {
            bool is16Bit = HasMultibyte(value);
            out1.WriteByte(is16Bit ? 0x01 : 0x00);
            if (is16Bit)
            {
                PutUnicodeLE(value, out1);
            }
            else
            {
                PutCompressedUnicode(value, out1);
            }
        }
        /// <summary>
        /// Gets the number of bytes that would be written by WriteUnicodeString(LittleEndianOutput, String)
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static int GetEncodedSize(String value)
        {
            int result = 2 + 1;
            result += value.Length * (StringUtil.HasMultibyte(value) ? 2 : 1);
            return result;
        }

        /// <summary>
        /// Checks to see if a given String needs to be represented as Unicode
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>
        /// 	<c>true</c> if string needs Unicode to be represented.; otherwise, <c>false</c>.
        /// </returns>
        ///<remarks>Tony Qu change the logic</remarks>
        public static bool IsUnicodeString(String value)
        {
            return !value.Equals(ISO_8859_1.GetString(ISO_8859_1.GetBytes(value)));
        }
        /// <summary> 
        /// Encodes non-US-ASCII characters in a string, good for encoding file names for download 
        /// http://www.acriticsreview.com/List.aspx?listid=42
        /// </summary> 
        /// <param name="s"></param> 
        /// <returns></returns> 
        public static string ToHexString(string s)
        {
            char[] chars = s.ToCharArray();
            StringBuilder builder = new StringBuilder();
            for (int index = 0; index < chars.Length; index++)
            {
                bool needToEncode = NeedToEncode(chars[index]);
                if (needToEncode)
                {
                    string encodedString = ToHexString(chars[index]);
                    builder.Append(encodedString);
                }
                else
                {
                    builder.Append(chars[index]);
                }
            }

            return builder.ToString();
        }
        /// <summary> 
        /// Encodes a non-US-ASCII character. 
        /// </summary> 
        /// <param name="chr"></param> 
        /// <returns></returns> 
        public static string ToHexString(char chr)
        {
            //UTF8Encoding utf8 = new UTF8Encoding();
            //byte[] encodedBytes = utf8.GetBytes(((int)chr).ToString());
            //StringBuilder builder = new StringBuilder();
            //for (int index = 0; index < encodedBytes.Length; index++)
            //{
            //    builder.AppendFormat("{0}", Convert.ToString(encodedBytes[index], 16));
            //}

            //return builder.ToString();

            return Convert.ToString((int)chr, 16);
        }

        /// <summary> 
        /// Encodes a non-US-ASCII character. 
        /// </summary> 
        /// <param name="chr"></param> 
        /// <returns></returns> 
        public static string ToHexString(short chr)
        {
            return ToHexString((char)chr);
        }

        /// <summary> 
        /// Encodes a non-US-ASCII character. 
        /// </summary> 
        /// <param name="chr"></param> 
        /// <returns></returns> 
        public static string ToHexString(int chr)
        {
            return ToHexString((char)chr);
        }
        /// <summary> 
        /// Encodes a non-US-ASCII character. 
        /// </summary> 
        /// <param name="chr"></param> 
        /// <returns></returns> 
        public static string ToHexString(long chr)
        {
            return ToHexString((char)chr);
        }
        /// <summary> 
        /// Determines if the character needs to be encoded. 
        /// http://www.acriticsreview.com/List.aspx?listid=42
        /// </summary> 
        /// <param name="chr"></param> 
        /// <returns></returns> 
        private static bool NeedToEncode(char chr)
        {
            string reservedChars = "$-_.+!*'(),@=&";

            if (chr > 127)
                return true;
            if (char.IsLetterOrDigit(chr) || reservedChars.IndexOf(chr) >= 0)
                return false;

            return true;
        }

        /**
        * Some strings may contain encoded characters of the unicode private use area.
        * Currently the characters of the symbol fonts are mapped to the corresponding
        * characters in the normal unicode range. 
        *
        * @param string the original string 
        * @return the string with mapped characters
        * 
        * @see <a href="http://www.alanwood.net/unicode/private_use_area.html#symbol">Private Use Area (symbol)</a>
        * @see <a href="http://www.alanwood.net/demos/symbol.html">Symbol font - Unicode alternatives for Greek and special characters in HTML</a>
        */
        public static String mapMsCodepointString(String string1)
        {
            if (string1 == null || "".Equals(string1)) return string1;
            InitMsCodepointMap();

            StringBuilder sb = new StringBuilder();
            int length = string1.Length;
            //char[] stringChars = string1.ToCharArray();
            for (int offset = 0; offset < length; )
            {
                
                int msCodepoint = char.ConvertToUtf32(string1, offset);//codePointAt(stringChars, offset, string1.Length);
                int uniCodepoint = msCodepointToUnicode[(msCodepoint)];
                sb.Append(Char.ConvertFromUtf32(uniCodepoint == null ? msCodepoint : uniCodepoint));
                offset += CharCount(msCodepoint);
            }

            return sb.ToString();
        }
        /**
         * The minimum value of a
         * <a href="http://www.unicode.org/glossary/#high_surrogate_code_unit">
         * Unicode high-surrogate code unit</a>
         * in the UTF-16 encoding, constant {@code '\u005CuD800'}.
         * A high-surrogate is also known as a <i>leading-surrogate</i>.
         *
         * @since 1.5
         */
        public static char MIN_HIGH_SURROGATE = '\uD800';

        /**
         * The maximum value of a
         * <a href="http://www.unicode.org/glossary/#high_surrogate_code_unit">
         * Unicode high-surrogate code unit</a>
         * in the UTF-16 encoding, constant {@code '\u005CuDBFF'}.
         * A high-surrogate is also known as a <i>leading-surrogate</i>.
         *
         * @since 1.5
         */
        public static char MAX_HIGH_SURROGATE = '\uDBFF';
            /**
         * The minimum value of a
         * <a href="http://www.unicode.org/glossary/#low_surrogate_code_unit">
         * Unicode low-surrogate code unit</a>
         * in the UTF-16 encoding, constant {@code '\u005CuDC00'}.
         * A low-surrogate is also known as a <i>trailing-surrogate</i>.
         *
         * @since 1.5
         */
        public static char MIN_LOW_SURROGATE  = '\uDC00';

        /**
         * The maximum value of a
         * <a href="http://www.unicode.org/glossary/#low_surrogate_code_unit">
         * Unicode low-surrogate code unit</a>
         * in the UTF-16 encoding, constant {@code '\u005CuDFFF'}.
         * A low-surrogate is also known as a <i>trailing-surrogate</i>.
         *
         * @since 1.5
         */
        public static char MAX_LOW_SURROGATE  = '\uDFFF';
            /**
         * Converts the specified surrogate pair to its supplementary code
         * point value. This method does not validate the specified
         * surrogate pair. The caller must validate it using {@link
         * #isSurrogatePair(char, char) isSurrogatePair} if necessary.
         *
         * @param  high the high-surrogate code unit
         * @param  low the low-surrogate code unit
         * @return the supplementary code point composed from the
         *         specified surrogate pair.
         * @since  1.5
         */
        public static int toCodePoint(char high, char low) {
            // Optimized form of:
            // return ((high - MIN_HIGH_SURROGATE) << 10)
            //         + (low - MIN_LOW_SURROGATE)
            //         + MIN_SUPPLEMENTARY_CODE_POINT;
            return ((high << 10) + low) + (MIN_SUPPLEMENTARY_CODE_POINT
                                           - (MIN_HIGH_SURROGATE << 10)
                                           - MIN_LOW_SURROGATE);
        }
        // throws ArrayIndexOutOfBoundsException if index out of bounds
        static int codePointAt(char[] a, int index, int limit)
        {
            char c1 = a[index];
            if (char.IsHighSurrogate(c1) && ++index < limit)
            {
                char c2 = a[index];
                if (char.IsLowSurrogate(c2))
                {
                    return toCodePoint(c1, c2);
                }
            }
            return c1;
        }
        public const int MIN_SUPPLEMENTARY_CODE_POINT = 0x010000;
        /**
         * Determines the number of {@code char} values needed to
         * represent the specified character (Unicode code point). If the
         * specified character is equal to or greater than 0x10000, then
         * the method returns 2. Otherwise, the method returns 1.
         *
         * This method doesn't validate the specified character to be a
         * valid Unicode code point. The caller must validate the
         * character value using {@link #isValidCodePoint(int) isValidCodePoint}
         * if necessary.
         *
         * @param   codePoint the character (Unicode code point) to be tested.
         * @return  2 if the character is a valid supplementary character; 1 otherwise.
         * @see     Character#isSupplementaryCodePoint(int)
         * @since   1.5
         */
        public static int CharCount(int codePoint)
        {
            return codePoint >= MIN_SUPPLEMENTARY_CODE_POINT ? 2 : 1;
        }
        public static void mapMsCodepoint(int msCodepoint, int unicodeCodepoint)
        {
            InitMsCodepointMap();
            msCodepointToUnicode.Add(msCodepoint, unicodeCodepoint);
        }

        private static void InitMsCodepointMap()
        {
            if (msCodepointToUnicode != null) return;
            msCodepointToUnicode = new Dictionary<int, int>();
            int i = 0xF020;
            foreach (int ch in symbolMap_f020)
            {
                msCodepointToUnicode.Add(i++, ch);
            }
            i = 0xf0a0;
            foreach (int ch in symbolMap_f0a0)
            {
                msCodepointToUnicode.Add(i++, ch);
            }
        }

        private static int[] symbolMap_f020 = {
       ' ', // 0xf020 space
       '!', // 0xf021 exclam
       8704, // 0xf022 universal
       '#', // 0xf023 numbersign
       8707, // 0xf024 existential
       '%', // 0xf025 percent
       '&', // 0xf026 ampersand
       8717, // 0xf027 suchthat
       '(', // 0xf028 parenleft
       ')', // 0xf029 parentright
       8727, // 0xf02a asteriskmath
       '+', // 0xf02b plus
       ',', // 0xf02c comma
       8722, // 0xf02d minus sign (long -)
       '.', // 0xf02e period
       '/', // 0xf02f slash
       '0', // 0xf030 0
       '1', // 0xf031 1
       '2', // 0xf032 2
       '3', // 0xf033 3
       '4', // 0xf034 4
       '5', // 0xf035 5
       '6', // 0xf036 6
       '7', // 0xf037 7
       '8', // 0xf038 8
       '9', // 0xf039 9
       ':', // 0xf03a colon
       ';', // 0xf03b semicolon
       '<', // 0xf03c less
       '=', // 0xf03d equal
       '>', // 0xf03e greater
       '?', // 0xf03f question
       8773, // 0xf040 congruent
       913, // 0xf041 alpha (upper)
       914, // 0xf042 beta (upper)
       935, // 0xf043 chi (upper)
       916, // 0xf044 delta (upper)
       917, // 0xf045 epsilon (upper)
       934, // 0xf046 phi (upper)
       915, // 0xf047 gamma (upper)
       919, // 0xf048 eta (upper)
       921, // 0xf049 iota (upper)
       977, // 0xf04a theta1 (lower)
       922, // 0xf04b kappa (upper)
       923, // 0xf04c lambda (upper)
       924, // 0xf04d mu (upper)
       925, // 0xf04e nu (upper)
       927, // 0xf04f omicron (upper)
       928, // 0xf050 pi (upper)
       920, // 0xf051 theta (upper)
       929, // 0xf052 rho (upper)
       931, // 0xf053 sigma (upper)
       932, // 0xf054 tau (upper)
       933, // 0xf055 upsilon (upper)
       962, // 0xf056 simga1 (lower)
       937, // 0xf057 omega (upper)
       926, // 0xf058 xi (upper)
       936, // 0xf059 psi (upper)
       918, // 0xf05a zeta (upper)
       '[', // 0xf05b bracketleft
       8765, // 0xf05c therefore
       ']', // 0xf05d bracketright
       8869, // 0xf05e perpendicular
       '_', // 0xf05f underscore
       ' ', // 0xf060 radicalex (doesn't exist in unicode)
       945, // 0xf061 alpha (lower)
       946, // 0xf062 beta (lower)
       967, // 0xf063 chi (lower)
       948, // 0xf064 delta (lower)
       949, // 0xf065 epsilon (lower)
       966, // 0xf066 phi (lower)
       947, // 0xf067 gamma (lower)
       951, // 0xf068 eta (lower)
       953, // 0xf069 iota (lower)
       981, // 0xf06a phi1 (lower)
       954, // 0xf06b kappa (lower)
       955, // 0xf06c lambda (lower)
       956, // 0xf06d mu (lower)
       957, // 0xf06e nu (lower)
       959, // 0xf06f omnicron (lower)
       960, // 0xf070 pi (lower)
       952, // 0xf071 theta (lower)
       961, // 0xf072 rho (lower)
       963, // 0xf073 sigma (lower)
       964, // 0xf074 tau (lower)
       965, // 0xf075 upsilon (lower)
       982, // 0xf076 piv (lower)
       969, // 0xf077 omega (lower)
       958, // 0xf078 xi (lower)
       968, // 0xf079 psi (lower)
       950, // 0xf07a zeta (lower)
       '{', // 0xf07b braceleft
       '|', // 0xf07c bar
       '}', // 0xf07d braceright
       8764, // 0xf07e similar '~'
       ' ', // 0xf07f not defined
   };

        private static int[] symbolMap_f0a0 = {
       8364, // 0xf0a0 not defined / euro symbol
       978, // 0xf0a1 upsilon1 (upper)
       8242, // 0xf0a2 minute
       8804, // 0xf0a3 lessequal
       8260, // 0xf0a4 fraction
       8734, // 0xf0a5 infInity
       402, // 0xf0a6 florin
       9827, // 0xf0a7 club
       9830, // 0xf0a8 diamond
       9829, // 0xf0a9 heart
       9824, // 0xf0aa spade
       8596, // 0xf0ab arrowboth
       8591, // 0xf0ac arrowleft
       8593, // 0xf0ad arrowup
       8594, // 0xf0ae arrowright
       8595, // 0xf0af arrowdown
       176, // 0xf0b0 degree
       177, // 0xf0b1 plusminus
       8243, // 0xf0b2 second
       8805, // 0xf0b3 greaterequal
       215, // 0xf0b4 multiply
       181, // 0xf0b5 proportional
       8706, // 0xf0b6 partialdiff
       8729, // 0xf0b7 bullet
       247, // 0xf0b8 divide
       8800, // 0xf0b9 notequal
       8801, // 0xf0ba equivalence
       8776, // 0xf0bb approxequal
       8230, // 0xf0bc ellipsis
       9168, // 0xf0bd arrowvertex
       9135, // 0xf0be arrowhorizex
       8629, // 0xf0bf carriagereturn
       8501, // 0xf0c0 aleph
       8475, // 0xf0c1 Ifraktur
       8476, // 0xf0c2 Rfraktur
       8472, // 0xf0c3 weierstrass
       8855, // 0xf0c4 circlemultiply
       8853, // 0xf0c5 circleplus
       8709, // 0xf0c6 emptyset
       8745, // 0xf0c7 intersection
       8746, // 0xf0c8 union
       8835, // 0xf0c9 propersuperset
       8839, // 0xf0ca reflexsuperset
       8836, // 0xf0cb notsubset
       8834, // 0xf0cc propersubset
       8838, // 0xf0cd reflexsubset
       8712, // 0xf0ce element
       8713, // 0xf0cf notelement
       8736, // 0xf0d0 angle
       8711, // 0xf0d1 gradient
       174, // 0xf0d2 registerserif
       169, // 0xf0d3 copyrightserif
       8482, // 0xf0d4 trademarkserif
       8719, // 0xf0d5 product
       8730, // 0xf0d6 radical
       8901, // 0xf0d7 dotmath
       172, // 0xf0d8 logicalnot
       8743, // 0xf0d9 logicaland
       8744, // 0xf0da logicalor
       8660, // 0xf0db arrowdblboth
       8656, // 0xf0dc arrowdblleft
       8657, // 0xf0dd arrowdblup
       8658, // 0xf0de arrowdblright
       8659, // 0xf0df arrowdbldown
       9674, // 0xf0e0 lozenge
       9001, // 0xf0e1 angleleft
       174, // 0xf0e2 registersans
       169, // 0xf0e3 copyrightsans
       8482, // 0xf0e4 trademarksans
       8721, // 0xf0e5 summation
       9115, // 0xf0e6 parenlefttp
       9116, // 0xf0e7 parenleftex
       9117, // 0xf0e8 parenleftbt
       9121, // 0xf0e9 bracketlefttp
       9122, // 0xf0ea bracketleftex
       9123, // 0xf0eb bracketleftbt
       9127, // 0xf0ec bracelefttp
       9128, // 0xf0ed braceleftmid
       9129, // 0xf0ee braceleftbt
       9130, // 0xf0ef braceex
       ' ', // 0xf0f0 not defined
       9002, // 0xf0f1 angleright
       8747, // 0xf0f2 integral
       8992, // 0xf0f3 integraltp
       9134, // 0xf0f4 integralex
       8993, // 0xf0f5 integralbt
       9118, // 0xf0f6 parenrighttp
       9119, // 0xf0f7 parenrightex
       9120, // 0xf0f8 parenrightbt
       9124, // 0xf0f9 bracketrighttp
       9125, // 0xf0fa bracketrightex
       9126, // 0xf0fb bracketrightbt
       9131, // 0xf0fc bracerighttp
       9132, // 0xf0fd bracerightmid
       9133, // 0xf0fe bracerightbt
       ' ', // 0xf0ff not defined
   };

    }
}

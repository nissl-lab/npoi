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
using System.IO;
using System.Globalization;


namespace NPOI.Util
{
    /// <summary>
    /// dump data in hexadecimal format; derived from a HexDump utility I
    /// wrote in June 2001.
    /// @author Marc Johnson
    /// @author Glen Stampoultzis  (glens at apache.org)
    /// </summary>
    public class HexDump
    {

        private static readonly char[] _hexcodes = new char[] { 
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 
        '\0', '\0'
     };
        private static readonly int[] _shifts = new int[] { 60, 56, 52, 48, 44, 40, 36, 32, 28, 24, 20, 16, 12, 8, 4, 0 };
        public static readonly string EOL = Environment.NewLine;

        private HexDump()
        {
        }

        private static string Dump(byte value)
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Length = 0;
            for (int i = 0; i < 2; i++)
            {
                buffer.Append(_hexcodes[(value >> _shifts[i + 6]) & 15]);
            }
            return buffer.ToString();
        }

        private static string Dump(long value)
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Length = 0;
            for (int i = 0; i < 8; i++)
            {
                buffer.Append(_hexcodes[((int)(value >> _shifts[i + _shifts.Length - 8])) & 15]);
            }
            return buffer.ToString();
        }

        public static string Dump(byte[] data, long offset, int index)
        {

            if ((index < 0) || (index >= data.Length))
            {
                string message = string.Format(CultureInfo.InvariantCulture, "illegal index: {0} into array of length {1}", index, data.Length);
                throw new IndexOutOfRangeException(message);
            }
            long display_offset = offset + index;
            StringBuilder buffer = new StringBuilder(0x4a);
            for (int i = index; i < data.Length; i += 16)
            {
                int chars_read = data.Length - i;

                if (chars_read > 16)
                {
                    chars_read = 16;
                }
                buffer.Append(Dump(display_offset)).Append(' ');
                for (int j = 0; j < 16; j++)
                {
                    if (j < chars_read)
                    {
                        buffer.Append(Dump(data[j + i]));
                    }
                    else
                    {
                        buffer.Append("  ");
                    }
                    buffer.Append(' ');
                }
                for (int k = 0; k < chars_read; k++)
                {
                    if ((data[k + i] >= ' ') && (data[k + i] < 127))
                    {
                        buffer.Append((char)data[k + i]);
                    }
                    else
                    {
                        buffer.Append('.');
                    }
                }
                buffer.Append(EOL);
                display_offset += chars_read;
            }
            return buffer.ToString();
        }

        public static void Dump(byte[] data, long offset, Stream stream, int index)
        {
            //lock (typeof(HexDump))
            {
                Dump(data, offset, stream, index, data.Length - index);
            }
        }

        public static void Dump(Stream inStream, int start, int bytesToDump)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                if (bytesToDump == -1)
                {
                    int c = inStream.ReadByte();
                    while (c != -1)
                    {
                        stream.WriteByte((byte)c);
                        c = inStream.ReadByte();
                    }
                }
                else
                {
                    int bytesRemaining = bytesToDump;
                    while (bytesRemaining-- > 0)
                    {
                        int c = inStream.ReadByte();
                        if (c == -1)
                            break;
                        else
                            stream.WriteByte((byte)c);
                    }
                }
                byte[] data = stream.ToArray();
                Dump(data, 0L, null, start, data.Length);
            }
        }

        public static void Dump(byte[] data, long offset, Stream stream, int index, int length)
        {
            //lock (typeof(HexDump))
            {
                if (data.Length == 0)
                {
                    byte[] info = Encoding.UTF8.GetBytes(string.Format(CultureInfo.InvariantCulture, "No Data{0}", EOL));
                    if (stream != null)
                    {
                        //Console.Write(info);
                        stream.Write(info, 0, info.Length);
                        stream.Flush();
                    }
                    return;
                }
                if ((index < 0) || (index >= data.Length))
                {
                    string message = string.Format(CultureInfo.InvariantCulture, "illegal index: {0} into array of length {1}", index, data.Length);
                    throw new IndexOutOfRangeException(message);
                }
                if (data.Length != 0)
                {

                    long display_offset = offset + index;
                    StringBuilder buffer = new StringBuilder(74);
                    int data_length = Math.Min(data.Length, index + length);
                    for (int i = index; i < data_length; i += 16)
                    {
                        int chars_read = data_length - i;
                        if (chars_read > 16)
                        {
                            chars_read = 16;
                        }
                        buffer.Append(Dump(display_offset)).Append(' ');
                        for (int j = 0; j < 16; j++)
                        {
                            if (j < chars_read)
                            {
                                buffer.Append(Dump(data[j + i]));
                            }
                            else
                            {
                                buffer.Append("  ");
                            }
                            buffer.Append(' ');
                        }
                        for (int k = 0; k < chars_read; k++)
                        {
                            if ((data[k + i] >= ' ') && (data[k + i] < 127))
                            {
                                buffer.Append((char)data[k + i]);
                            }
                            else
                            {
                                buffer.Append('.');
                            }
                        }
                        buffer.Append(EOL);
                        byte[] bytes = Encoding.UTF8.GetBytes(buffer.ToString());
                        if (stream != null)
                        {
                            //Console.Write(buffer.ToString());
                            stream.Write(bytes, 0, bytes.Length);
                            // Console.Write(Encoding.UTF8.GetString(bytes));
                            stream.Flush();
                        }
                        buffer.Length = 0;
                        display_offset += chars_read;
                    }
                }
            }
        }


        /**
         * Dumps <code>bytesToDump</code> bytes to an output stream.
         *
         * @param in          The stream to read from
         * @param out         The output stream
         * @param start       The index to use as the starting position for the left hand side label
         * @param bytesToDump The number of bytes to output.  Use -1 to read until the end of file.
         */
        public static void Dump(Stream in1, Stream out1, int start, int bytesToDump ) 
        {
            MemoryStream buf = new MemoryStream();
            if (bytesToDump == -1)
            {
                int c = in1.ReadByte();
                while (c != -1)
                {
                    buf.WriteByte((byte)c);
                    c = in1.ReadByte();
                }
            }
            else
            {
                int bytesRemaining = bytesToDump;
                while (bytesRemaining-- > 0)
                {
                    int c = in1.ReadByte();
                    if (c == -1) {
                        break;
                    }
                    buf.WriteByte((byte)c);
                }
            }

            byte[] data = buf.ToArray();
            Dump(data, 0, out1, start, data.Length);
        }

        /// <summary>
        /// Shorts to hex.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>char array of 2 (zero padded) uppercase hex chars and prefixed with '0x'</returns>
        public static char[] ShortToHex(int value)
        {
            return ToHexChars(value, 2);
        }
        /// <summary>
        /// Bytes to hex.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>char array of 1 (zero padded) uppercase hex chars and prefixed with '0x'</returns>
        public static char[] ByteToHex(int value)
        {
            return ToHexChars(value, 1);
        }

        /// <summary>
        /// Ints to hex.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>char array of 4 (zero padded) uppercase hex chars and prefixed with '0x'</returns>
        public static char[] IntToHex(int value)
        {
            return ToHexChars(value, 4);
        }
        /// <summary>
        /// char array of 4 (zero padded) uppercase hex chars and prefixed with '0x'
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>char array of 4 (zero padded) uppercase hex chars and prefixed with '0x'</returns>
        public static char[] LongToHex(long value)
        {
            return ToHexChars(value, 8);
        }
        /// <summary>
        /// Toes the hex chars.
        /// </summary>
        /// <param name="pValue">The p value.</param>
        /// <param name="nBytes">The n bytes.</param>
        /// <returns>char array of uppercase hex chars, zero padded and prefixed with '0x'</returns>
        private static char[] ToHexChars(long pValue, int nBytes)
        {
            int charPos = 2 + nBytes * 2;
            // The return type is char array because most callers will probably append the value to a
            // StringBuffer, or write it to a Stream / Writer so there is no need to create a String;
            char[] result = new char[charPos];

            long value = pValue;
            do
            {
                result[--charPos] = _hexcodes[(int)(value & 0x0F)];
                value >>= 4;
            } while (charPos > 1);

            // Prefix added to avoid ambiguity
            result[0] = '0';
            result[1] = 'x';
            return result;
        }

        public static string ToHex(byte value)
        {
            return ToHex((long)value, 2);
        }

        public static string ToHex(short value)
        {
            return ToHex((long)value, 4);
        }

        public static string ToHex(int value)
        {
            return ToHex((long)value, 8);
        }

        public static string ToHex(long value)
        {
            return ToHex(value, 16);
        }

        public static string ToHex(byte[] value)
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append('[');
            for (int i = 0; i < value.Length; i++)
            {
                if (i > 0)
                {
                    buffer.Append(", ");
                }
                buffer.Append(ToHex(value[i]));
            }
            buffer.Append(']');
            return buffer.ToString();
        }

        public static string ToHex(short[] value)
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append('[');
            for (int i = 0; i < value.Length; i++)
            {
                if (i > 0)
                {
                    buffer.Append(", ");
                }
                buffer.Append(ToHex(value[i]));
                
            }
            buffer.Append(']');
            return buffer.ToString();
        }

        private static string ToHex(long value, int digits)
        {
            StringBuilder buffer = new StringBuilder(digits);
            for (int i = 0; i < digits; i++)
            {
                buffer.Append(_hexcodes[(int)((value >> _shifts[i + (16 - digits)]) & 15L)]);
            }
            return buffer.ToString();
        }

        public static String ToHex(byte[] value, int bytesPerLine)
        {
            int digits = value.Length == 0 ? 0 : (int)Math.Round(Math.Log(value.Length) / Math.Log(10) + 0.50000001);
            StringBuilder formatString = new StringBuilder();
            for (int i = 0; i < digits; i++)
                formatString.Append('0');
            formatString.Append(": ");
            StringBuilder retVal = new StringBuilder();
            retVal.Append(((double)0).ToString(formatString.ToString(), CultureInfo.InvariantCulture));
            if (value.Length == 0)
                retVal.Append("0");
            int j = -1;
            for (int x = 0; x < value.Length; x++)
            {
                if (++j == bytesPerLine)
                {
                    retVal.Append('\n');
                    retVal.Append(((double)x).ToString(formatString.ToString(), CultureInfo.InvariantCulture));
                    j = 0;
                }
                else if (x > 0)
                {
                    retVal.Append(", ");
                }
                retVal.Append(ToHex(value[x]));
            }
            return retVal.ToString();
        }
    }

}

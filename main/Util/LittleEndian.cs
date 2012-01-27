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
 * Revision
 * Who          Date            Comment  
 * Tony Qu      2008.10.18      add methods to support ulong
 * ==============================================================*/

using System;
using System.IO;

namespace NPOI.Util
{
    /// <summary>
    /// a utility class for handling little-endian numbers, which the 80x86 world is
    /// replete with. The methods are all static, and input/output is from/to byte
    /// arrays, or from InputStreams.
    /// </summary>
    /// <remarks>
    /// @author     Marc Johnson (mjohnson at apache dot org)
    /// @author     Andrew Oliver (acoliver at apache dot org)
    /// </remarks>
    public class LittleEndian 
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LittleEndian"/> class.
        /// </summary>
        private LittleEndian()
        {
        }

        /// <summary>
        /// Copy a portion of a byte array
        /// </summary>
        /// <param name="data"> the original byte array</param>
        /// <param name="offset">Where to start copying from.</param>
        /// <param name="size">Number of bytes to copy.</param>
        /// <returns>The byteArray value</returns>
        ///<exception cref="IndexOutOfRangeException">
        ///if copying would cause access ofdata outside array bounds.
        ///</exception>
        public static byte[] GetByteArray(byte[] data, int offset, int size)
        {
            byte[] destinationArray = new byte[size];
            Array.Copy(data, offset, destinationArray, 0, size);
            return destinationArray;
        }

        public static double GetDouble(byte[] data)
        {
            return GetDouble(data, 0);
        }

        public static double GetDouble(byte[] data, int offset)
        {
            return BitConverter.Int64BitsToDouble(GetNumber(data, offset, LittleEndianConstants.DOUBLE_SIZE));
        }

        public static int GetInt(byte[] data)
        {
            return GetInt(data, 0);
        }

        public static int GetInt(byte[] data, int offset)
        {
            return (int)GetNumber(data, offset, LittleEndianConstants.INT_SIZE);
        }

        public static long GetLong(byte[] data)
        {
            return GetLong(data, 0);
        }

        public static long GetLong(byte[] data, int offset)
        {
            return GetNumber(data, offset, LittleEndianConstants.LONG_SIZE);
        }

        public static ulong GetULong(byte[] data)
        {
            return GetULong(data, 0);
        }

        public static ulong GetULong(byte[] data, int offset)
        {
            return BitConverter.ToUInt64(data, offset);
        }

        private static long GetNumber(byte[] data, int offset, int size)
        {
            long num = 0L;
            for (int i = (offset + size) - 1; i >= offset; i--)
            {
                num = num << 8;
                num |= 0xffL & data[i];
            }
            return num;
        }

        /// <summary>
        /// get a short value from a byte array
        /// </summary>
        /// <param name="data">a starting offset into the byte array</param>
        /// <returns>the short (16-bit) value</returns>
        public static short GetShort(byte[] data)
        {
            return GetShort(data, 0);
        }

        /// <summary>
        /// get a short value from a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the short (16-bit) value</returns>
        public static short GetShort(byte[] data, int offset)
        {
            return (short)GetNumber(data, offset, LittleEndianConstants.SHORT_SIZE);
        }

        /// <summary>
        /// Gets the U int.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <returns></returns>
        public static uint GetUInt(byte[] data)
        {
            return GetUInt(data, 0);
        }

        /// <summary>
        /// Gets the U int.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns></returns>
        public static uint GetUInt(byte[] data, int offset)
        {
            uint num = (uint)GetNumber(data, offset, LittleEndianConstants.INT_SIZE);
            if (num < 0)
            {
                return (uint)(0x100000000L + num);
            }
            return num;
        }

        /// <summary>
        /// Gets the unsigned byte.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <returns></returns>
        public static byte GetUByte(byte[] data)
        {
            return GetUByte(data, 0);
        }

        /// <summary>
        /// Gets the unsigned byte.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns></returns>
        public static byte GetUByte(byte[] data, int offset)
        {
            return (byte)GetNumber(data, offset, LittleEndianConstants.BYTE_SIZE);
        }

        /// <summary>
        /// get a short value from a byte array
        /// </summary>
        /// <param name="data">the unsigned short (16-bit) value in an integer</param>
        /// <returns></returns>
        public static ushort GetUShort(byte[] data)
        {
            return GetUShort(data, 0);
        }

        /// <summary>
        /// get an unsigned short value from a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the unsigned short (16-bit) value in an integer</returns>
        public static ushort GetUShort(byte[] data, int offset)
        {
            short num = (short)GetNumber(data, offset, LittleEndianConstants.SHORT_SIZE);
            if (num < 0)
            {
                return (ushort)(0x10000 + num);
            }
            return (ushort)num;
        }

        /// <summary>
        /// Puts the double.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutDouble(byte[] data, double value)
        {
            PutDouble(data, 0, value);
        }

        /// <summary>
        /// Puts the double.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutDouble(byte[] data, int offset, double value)
        {
            if (double.IsNaN(value))
            {
                PutNumber( data, offset, -276939487313920L, LittleEndianConstants.DOUBLE_SIZE);
            }
            else
            {
                PutNumber( data, offset, BitConverter.DoubleToInt64Bits(value), LittleEndianConstants.DOUBLE_SIZE);
            }   
        }

        /// <summary>
        /// Puts the int.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutInt(byte[] data, int value)
        {
            PutInt(data, 0, value);
        }

        /// <summary>
        /// Puts the int.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutInt(byte[] data, int offset, int value)
        {
            PutNumber( data, offset,Convert.ToInt64(value), LittleEndianConstants.INT_SIZE);
        }

        /// <summary>
        /// Puts the uint.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutUInt(byte[] data, uint value)
        {
            PutUInt(data, 0, value);
        }

        /// <summary>
        /// Puts the uint.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutUInt(byte[] data, int offset, uint value)
        {
            PutNumber(data, offset, Convert.ToInt64(value), LittleEndianConstants.UINT_SIZE);
        }

        /// <summary>
        /// Puts the long.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutLong(byte[] data, long value)
        {
            PutLong(data, 0, value);
        }

        /// <summary>
        /// Puts the long.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutLong(byte[] data, int offset, long value)
        {
            PutNumber( data, offset, value, LittleEndianConstants.LONG_SIZE);
        }

        /// <summary>
        /// Puts the long.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutULong(byte[] data, ulong value)
        {
            PutULong(data, 0, value);
        }

        /// <summary>
        /// Puts the ulong.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutULong(byte[] data, int offset, ulong value)
        {
            PutNumber(data, offset, value, LittleEndianConstants.ULONG_SIZE);
        }

        /// <summary>
        /// Puts the number.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        /// <param name="size">The size.</param>
        private static void PutNumber(byte[] data, int offset, long value, int size)
        {
            int limit = size + offset;
            long v = value;
            for (int i = offset; i < limit; i++)
            {
                data[i] = (byte)(v & 0xffL);
                v >>= 8;
            }
        }

        /// <summary>
        /// Puts the number.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        /// <param name="size">The size.</param>
        private static void PutNumber(byte[] data, int offset, ulong value, int size)
        {
            int limit = size + offset;
            ulong v = value;
            for (int i = offset; i < limit; i++)
            {
                data[i] = (byte)(v & 0xffL);
                v >>= 8;
            }
        }

        /// <summary>
        /// Puts the short.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutShort(byte[] data, short value)
        {
            PutShort(data, 0, value);
        }

        /// <summary>
        /// Puts the short.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutShort(byte[] data, int offset, short value)
        {
            PutNumber( data, offset, Convert.ToInt64(value), LittleEndianConstants.SHORT_SIZE);
        }

        /// <summary>
        /// Puts the short array.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutShortArray(byte[] data, int offset, short[] value)
        {
            PutNumber( data, offset, Convert.ToInt64(value.Length), LittleEndianConstants.SHORT_SIZE);
            for (int i = 0; i < value.Length; i++)
            {
                PutNumber( data, (offset + 2) + (i * 2), Convert.ToInt64(value[i]), LittleEndianConstants.SHORT_SIZE);
            }
        }

        /// <summary>
        /// Added for consistency with other put~() methods
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutByte(byte[] data, int offset, int value)
        {
            PutNumber(data, offset, value, LittleEndianConstants.BYTE_SIZE);
        }

        /// <summary>
        /// Puts the U short.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutUShort(byte[] data, int value)
        {
            PutNumber(data, 0, Convert.ToInt64(value), LittleEndianConstants.SHORT_SIZE);
        }

        /// <summary>
        /// Puts the U short.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutUShort(byte[] data, int offset, int value)
        {
            PutNumber( data, offset, Convert.ToInt64(value), LittleEndianConstants.SHORT_SIZE);
        }

        /// <summary>
        /// Reads from stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="size">The size.</param>
        /// <returns></returns>
        public static byte[] ReadFromStream(Stream stream, int size)
        {
            byte[] buffer = new byte[size];
            int num = stream.Read(buffer, 0, buffer.Length);
            if (num == 0)
            {
                for (int i = 0; i < buffer.Length; i++)
                {
                    buffer[i] = 0;
                }
                return buffer;
            }
            if (num != size)
            {
                throw new BufferUnderrunException();
            }
            return buffer;
        }

        /// <summary>
        /// Reads the int.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public static int ReadInt(Stream stream)
        {
            return GetInt(ReadFromStream(stream, LittleEndianConstants.INT_SIZE));
        }

        /// <summary>
        /// Reads the long.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public static long ReadLong(Stream stream)
        {
            return GetLong(ReadFromStream(stream, LittleEndianConstants.LONG_SIZE));
        }

        /// <summary>
        /// Reads the long.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public static ulong ReadULong(Stream stream)
        {
            return GetULong(ReadFromStream(stream, LittleEndianConstants.LONG_SIZE));
        }

        /// <summary>
        /// Reads the short.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public static short ReadShort(Stream stream)
        {
            return GetShort(ReadFromStream(stream, LittleEndianConstants.SHORT_SIZE));
        }

        /// <summary>
        /// Us the byte to int.
        /// </summary>
        /// <param name="b">The b.</param>
        /// <returns></returns>
        public static int UByteToInt(byte b)
        {
            return (((b & 0x80) != 0) ? ((b & 0x7f) + 0x80) : b);
        }
    }
    // Nested Types
    [Serializable]
    public class BufferUnderrunException : IOException
    {
        // Methods
        internal BufferUnderrunException()
            : base("buffer underrun")
        {
        }
    }
}

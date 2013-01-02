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
    public class LittleEndian : LittleEndianConsts
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LittleEndian"/> class.
        /// </summary>
        private LittleEndian()
        {
            // no instances of this class
        }

        /// <summary>
        /// get a short value from a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the short (16-bit) value</returns>
        public static short GetShort(byte[] data, int offset)
        {
            int b0 = data[offset] & 0xFF;
            int b1 = data[offset + 1] & 0xFF;
            return (short)((b1 << 8) + (b0 << 0));
            //return (short)GetNumber(data, offset, LittleEndianConsts.SHORT_SIZE);
        }

        /// <summary>
        /// get an unsigned short value from a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the unsigned short (16-bit) value in an integer</returns>
        public static int GetUShort(byte[] data, int offset)
        {
            //short num = (short)GetNumber(data, offset, LittleEndianConsts.SHORT_SIZE);
            //if (num < 0)
            //{
            //    return (ushort)(0x10000 + num);
            //}
            //return (ushort)num;
            int b0 = data[offset] & 0xFF;
            int b1 = data[offset + 1] & 0xFF;
            return (b1 << 8) + (b0 << 0);
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
        /// <param name="data">the unsigned short (16-bit) value in an integer</param>
        /// <returns></returns>
        public static int GetUShort(byte[] data)
        {
            return GetUShort(data, 0);
        }

        /// <summary>
        /// get an int value from a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the int (32-bit) value</returns>
        public static int GetInt(byte[] data, int offset)
        {
            //return (int)GetNumber(data, offset, LittleEndianConsts.INT_SIZE);
            int i = offset;
            int b0 = data[i++] & 0xFF;
            int b1 = data[i++] & 0xFF;
            int b2 = data[i++] & 0xFF;
            int b3 = data[i++] & 0xFF;
            return (b3 << 24) + (b2 << 16) + (b1 << 8) + (b0 << 0);
        }

        /// <summary>
        /// get an int value from the beginning of a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <returns>the int (32-bit) value</returns>
        public static int GetInt(byte[] data)
        {
            return GetInt(data, 0);
        }

        /// <summary>
        /// Gets the U int.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the unsigned int (32-bit) value in a long</returns>
        public static long GetUInt(byte[] data, int offset)
        {
            //uint num = (uint)GetNumber(data, offset, LittleEndianConsts.INT_SIZE);
            //if (num < 0)
            //{
            //    return (uint)(0x100000000L + num);
            //}
            //return num;
            long retNum = GetInt(data, offset);
            return retNum & 0x00FFFFFFFFL;
        }

        /// <summary>
        /// Gets the U int.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <returns>the unsigned int (32-bit) value in a long</returns>
        public static long GetUInt(byte[] data)
        {
            return GetUInt(data, 0);
        }

        /// <summary>
        /// get a long value from a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the long (64-bit) value</returns>
        public static long GetLong(byte[] data, int offset)
        {
            //return GetNumber(data, offset, LittleEndianConsts.LONG_SIZE);
            long result = 0;

            for (int j = offset + LONG_SIZE - 1; j >= offset; j--)
            {
                result <<= 8;
                result |= 0xffL & data[j];
            }
            return result;
        }

        /// <summary>
        /// get a double value from a byte array, reads it in little endian format
        /// then converts the resulting revolting IEEE 754 (curse them) floating
        /// point number to a c# double
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns>the double (64-bit) value</returns>
        public static double GetDouble(byte[] data, int offset)
        {
            return BitConverter.Int64BitsToDouble(GetLong(data, offset));
        }

        /// <summary>
        /// Puts the short.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutShort(byte[] data, int offset, short value)
        {
            int i = offset;
            data[i++] = (byte)((value >> 0) & 0xFF);
            data[i++] = (byte)((value >> 8) & 0xFF);
            //PutNumber(data, offset, Convert.ToInt64(value), LittleEndianConsts.SHORT_SIZE);
        }

        /// <summary>
        /// Added for consistency with other put~() methods
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutByte(byte[] data, int offset, int value)
        {
            data[offset] = (byte)value;
            //PutNumber(data, offset, value, LittleEndianConsts.BYTE_SIZE);
        }


        /// <summary>
        /// Puts the U short.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        public static void PutUShort(byte[] data, int offset, int value)
        {
            //PutNumber(data, offset, Convert.ToInt64(value), LittleEndianConsts.SHORT_SIZE);
            int i = offset;
            data[i++] = (byte)((value >> 0) & 0xFF);
            data[i++] = (byte)((value >> 8) & 0xFF);
        }


        /// <summary>
        /// put a short value into beginning of a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">the short (16-bit) value</param>
        [Obsolete]
        public static void PutShort(byte[] data, short value)
        {
            PutShort(data, 0, value);
        }

        /**
         * Put signed short into output stream
         * 
         * @param value
         *            the short (16-bit) value
         * @param outputStream
         *            output stream
         * @throws IOException
         *             if an I/O error occurs
         */
        public static void PutShort( Stream outputStream, short value )
        {
            outputStream.WriteByte((byte)((value >> 0) & 0xFF));
            outputStream.WriteByte((byte)((value >> 8) & 0xFF));
        }

        /// <summary>
        /// put an int value into a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">the int (32-bit) value</param>
        public static void PutInt(byte[] data, int offset, int value)
        {
            int i = offset;
            data[i++] = (byte)((value >> 0) & 0xFF);
            data[i++] = (byte)((value >> 8) & 0xFF);
            data[i++] = (byte)((value >> 16) & 0xFF);
            data[i++] = (byte)((value >> 24) & 0xFF);
        }

        /// <summary>
        /// put an int value into beginning of a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">the int (32-bit) value</param>
        [Obsolete]
        public static void PutInt(byte[] data, int value)
        {
            PutInt(data, 0, value);
        }

        /// <summary>
        /// Put int into output stream
        /// </summary>
        /// <param name="value">the int (32-bit) value</param>
        /// <param name="outputStream">output stream</param>
        public static void PutInt(int value, Stream outputStream)
        {
            outputStream.WriteByte((byte) ((value >> 0) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 8) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 16) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 24) & 0xFF));
        }

        /// <summary>
        /// put a long value into a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">the long (64-bit) value</param>
        public static void PutLong(byte[] data, int offset, long value)
        {
            //PutNumber(data, offset, value, LittleEndianConsts.LONG_SIZE);
            int limit = LONG_SIZE + offset;
            long v = value;

            for (int j = offset; j < limit; j++)
            {
                data[j] = (byte)(v & 0xFF);
                v >>= 8;
            }
        }


        /// <summary>
        /// put a double value into a byte array
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">the double (64-bit) value</param>
        public static void PutDouble(byte[] data, int offset, double value)
        {
            long lvalue = 0L;
            if (double.IsNaN(value))
            {
                lvalue = -276939487313920L;
                //PutNumber(data, offset, -276939487313920L, LittleEndianConsts.DOUBLE_SIZE);
            }
            else
            {
                lvalue = BitConverter.DoubleToInt64Bits(value);
                //PutNumber(data, offset, BitConverter.DoubleToInt64Bits(value), LittleEndianConsts.DOUBLE_SIZE);
            }
            PutLong(data, offset, lvalue);
        }

        /// <summary>
        /// Reads the short.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public static short ReadShort(Stream stream)
        {
            //return GetShort(ReadFromStream(stream, LittleEndianConsts.SHORT_SIZE));
            return (short)ReadUShort(stream);
        }

        public static int ReadUShort(Stream stream)
        {
            int ch1 = stream.ReadByte();
            int ch2 = stream.ReadByte();
            if ((ch1 | ch2) < 0)
            {
                throw new BufferUnderrunException();
            }
            return (ch2 << 8) + (ch1 << 0);
        }
        /// <summary>
        /// get an int value from an Stream
        /// </summary>
        /// <param name="stream">the Stream from which the int is to be read</param>
        /// <returns>the int (32-bit) value</returns>
        /// <exception cref="T:System.IO.IOException">will be propagated back to the caller</exception>
        /// <exception cref="T:NPOI.Util.BufferUnderrunException">if the stream cannot provide enough bytes</exception>
        public static int ReadInt(Stream stream)
        {
            //return GetInt(ReadFromStream(stream, LittleEndianConsts.INT_SIZE));
            int ch1 = stream.ReadByte();
            int ch2 = stream.ReadByte();
            int ch3 = stream.ReadByte();
            int ch4 = stream.ReadByte();
            if ((ch1 | ch2 | ch3 | ch4) < 0)
            {
                throw new BufferUnderrunException();
            }
            return (ch4 << 24) + (ch3 << 16) + (ch2 << 8) + (ch1 << 0);
        }

        /// <summary>
        /// get a long value from a Stream
        /// </summary>
        /// <param name="stream">the Stream from which the long is to be read</param>
        /// <returns>the long (64-bit) value</returns>
        /// <exception cref="T:System.IO.IOException">will be propagated back to the caller</exception>
        /// <exception cref="T:NPOI.Util.BufferUnderrunException">if the stream cannot provide enough bytes</exception>
        public static long ReadLong(Stream stream)
        {
            //return GetLong(ReadFromStream(stream, LittleEndianConsts.LONG_SIZE));
            int ch1 = stream.ReadByte();
            int ch2 = stream.ReadByte();
            int ch3 = stream.ReadByte();
            int ch4 = stream.ReadByte();
            int ch5 = stream.ReadByte();
            int ch6 = stream.ReadByte();
            int ch7 = stream.ReadByte();
            int ch8 = stream.ReadByte();
            if ((ch1 | ch2 | ch3 | ch4 | ch5 | ch6 | ch7 | ch8) < 0)
            {
                throw new BufferUnderrunException();
            }

            return
                ((long)ch8 << 56) +
                ((long)ch7 << 48) +
                ((long)ch6 << 40) +
                ((long)ch5 << 32) +
                ((long)ch4 << 24) + // cast to long to preserve bit 31 (sign bit for ints)
                      (ch3 << 16) +
                      (ch2 << 8) +
                      (ch1 << 0);
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

        /// <summary>
        /// get the unsigned value of a byte.
        /// </summary>
        /// <param name="data">the byte array.</param>
        /// <param name="offset">a starting offset into the byte array.</param>
        /// <returns>the unsigned value of the byte as a 32 bit integer</returns>
        [Obsolete]
        public static int GetUnsignedByte(byte[] data, int offset)
        {
            return data[offset] & 0xFF;
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

        //Are the following methods Obsoleted?
        [Obsolete]
        public static double GetDouble(byte[] data)
        {
            return GetDouble(data, 0);
        }
        [Obsolete]
        public static long GetLong(byte[] data)
        {
            return GetLong(data, 0);
        }


        [Obsolete]
        public static ulong GetULong(byte[] data)
        {
            return GetULong(data, 0);
        }
        [Obsolete]
        public static ulong GetULong(byte[] data, int offset)
        {
            return BitConverter.ToUInt64(data, offset);
        }
        //[Obsolete] - private method can not be used outside of this class and only obsolete methods are using this method.
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
        /// Gets the unsigned byte.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <returns></returns>
        public static short GetUByte(byte[] data)
        {
            return (short)(data[0] & 0xFF);
        }

        /// <summary>
        /// Gets the unsigned byte.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <returns></returns>
        public static short GetUByte(byte[] data, int offset)
        {
            return (short)(data[offset] & 0xFF);
        }
        

        

        /// <summary>
        /// Puts the double.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        [Obsolete]
        public static void PutDouble(byte[] data, double value)
        {
            PutDouble(data, 0, value);
        }

        /**
         * put a double value into a byte array
         * 
         * @param value
         *            the double (64-bit) value
         * @param outputStream
         *            output stream
         * @throws IOException
         *             if an I/O error occurs
         */

        public static void PutDouble(double value, Stream outputStream)
        {
            PutLong(BitConverter.DoubleToInt64Bits(value), outputStream);
        }



        /// <summary>
        /// Puts the uint.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        [Obsolete]
        public static void PutUInt(byte[] data, uint value)
        {
            PutUInt(data, 0, value);
        }

        /**
         * Put unsigned int into output stream
         * 
         * @param value
         *            the int (32-bit) value
         * @param outputStream
         *            output stream
         * @throws IOException
         *             if an I/O error occurs
         */
        public static void PutUInt( long value, Stream outputStream )
        {
            outputStream.WriteByte((byte)((value >> 0) & 0xFF));
            outputStream.WriteByte((byte)((value >> 8) & 0xFF));
            outputStream.WriteByte((byte)((value >> 16) & 0xFF));
            outputStream.WriteByte((byte)((value >> 24) & 0xFF));
        }

        /// <summary>
        /// Puts the uint.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        [Obsolete]
        public static void PutUInt(byte[] data, int offset, uint value)
        {
            PutNumber(data, offset, Convert.ToInt64(value), LittleEndianConsts.UINT_SIZE);
        }

        /// <summary>
        /// Puts the long.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        [Obsolete]
        public static void PutLong(byte[] data, long value)
        {
            PutLong(data, 0, value);
        }

        /**
         * Put long into output stream
         * 
         * @param value
         *            the long (64-bit) value
         * @param outputStream
         *            output stream
         * @throws IOException
         *             if an I/O error occurs
         */
        public static void PutLong( long value, Stream outputStream )
        {
            outputStream.WriteByte((byte) ((value >> 0) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 8) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 16) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 24) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 32) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 40) & 0xFF));
            outputStream.WriteByte((byte) ((value >> 48) & 0xFF));
            outputStream.WriteByte((byte)((value >> 56) & 0xFF));
        }

        /// <summary>
        /// Puts the long.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        [Obsolete]
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
        [Obsolete]
        public static void PutULong(byte[] data, int offset, ulong value)
        {
            PutNumber(data, offset, value, LittleEndianConsts.ULONG_SIZE);
        }

        /// <summary>
        /// Puts the number.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        /// <param name="size">The size.</param>
        //[Obsolete] - private method can not be used outside of this class and only obsolete methods are using this method.
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
        //[Obsolete] - private method can not be used outside of this class and only obsolete methods are using this method.
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
        /// Puts the short array.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="offset">a starting offset into the byte array</param>
        /// <param name="value">The value.</param>
        [Obsolete]
        public static void PutShortArray(byte[] data, int offset, short[] value)
        {
            PutNumber( data, offset, Convert.ToInt64(value.Length), LittleEndianConsts.SHORT_SIZE);
            for (int i = 0; i < value.Length; i++)
            {
                PutNumber( data, (offset + 2) + (i * 2), Convert.ToInt64(value[i]), LittleEndianConsts.SHORT_SIZE);
            }
        }

        

        /// <summary>
        /// Puts the U short.
        /// </summary>
        /// <param name="data">the byte array</param>
        /// <param name="value">The value.</param>
        [Obsolete]
        public static void PutUShort(byte[] data, int value)
        {
            PutNumber(data, 0, Convert.ToInt64(value), LittleEndianConsts.SHORT_SIZE);
        }
        /**
         * Put unsigned short into output stream
         * 
         * @param value
         *            the unsigned short (16-bit) value
         * @param outputStream
         *            output stream
         * @throws IOException
         *             if an I/O error occurs
         */
        public static void PutUShort( int value, Stream outputStream )
        {
            outputStream.WriteByte((byte)((value >> 0) & 0xFF));
            outputStream.WriteByte((byte)((value >> 8) & 0xFF));
        }

        /// <summary>
        /// Reads from stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="size">The size.</param>
        /// <returns></returns>
        [Obsolete]
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
        /// Reads the long.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        [Obsolete]
        public static ulong ReadULong(Stream stream)
        {
            //return GetULong(ReadFromStream(stream, LittleEndianConsts.LONG_SIZE));
            return BitConverter.ToUInt64(ReadFromStream(stream, LittleEndianConsts.LONG_SIZE), 0);
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
    [Serializable]
    public class BufferUnderflowException : RuntimeException
    {
        public BufferUnderflowException()
            : base("Buffer Underflow")
        {
        }
    }
}

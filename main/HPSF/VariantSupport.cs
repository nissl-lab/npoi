/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
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

namespace NPOI.HPSF
{
    using System;
    using System.Collections;
    using System.IO;
    using System.Text;
    using NPOI.Util;
    using System.Globalization;

    /// <summary>
    /// Supports Reading and writing of variant data.
    /// <strong>FIXME (3):</strong>
    ///  Reading and writing should be made more
    /// uniform than it is now. The following items should be resolved:
    /// Reading requires a Length parameter that is 4 byte greater than the
    /// actual data, because the variant type field is included.
    /// Reading Reads from a byte array while writing Writes To an byte array
    /// output stream.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2003-08-08
    /// </summary>
    public class VariantSupport : Variant
    {

        private static bool logUnsupportedTypes = false;

        /// <summary>
        /// Checks whether logging of unsupported variant types warning is turned
        /// on or off.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if logging is turned on; otherwise, <c>false</c>.
        /// </value>
        public static bool IsLogUnsupportedTypes
        {
            get
            {
                return logUnsupportedTypes;
            }
            set 
            { 
                logUnsupportedTypes = value; 
            }
        }



        /**
         * Keeps a list of the variant types an "unsupported" message has alReady
         * been issued for.
         */
        protected static IList unsupportedMessage;

        /// <summary>
        /// Writes a warning To System.err that a variant type Is
        /// unsupported by HPSF. Such a warning is written only once for each variant
        /// type. Log messages can be turned on or off by
        /// </summary>
        /// <param name="ex">The exception To log</param>
        public static void WriteUnsupportedTypeMessage
            (UnsupportedVariantTypeException ex)
        {
            if (IsLogUnsupportedTypes)
            {
                if (unsupportedMessage == null)
                    unsupportedMessage = new ArrayList();
                long vt = ex.VariantType;
                if (!unsupportedMessage.Contains(vt))
                {
                    Console.Error.WriteLine(ex.Message);
                    unsupportedMessage.Add(vt);
                }
            }
        }


        /**
         * HPSF is able To Read these {@link Variant} types.
         */
        static public int[] SUPPORTED_TYPES = { Variant.VT_EMPTY,
            Variant.VT_I2, Variant.VT_I4, Variant.VT_I8, Variant.VT_R8,
            Variant.VT_FILETIME, Variant.VT_LPSTR, Variant.VT_LPWSTR,
            Variant.VT_CF, Variant.VT_BOOL };



        /// <summary>
        /// Checks whether HPSF supports the specified variant type. Unsupported
        /// types should be implemented included in the {@link #SUPPORTED_TYPES}
        /// array.
        /// </summary>
        /// <param name="variantType">the variant type To check</param>
        /// <returns>
        /// 	<c>true</c> if HPFS supports this type,otherwise, <c>false</c>.
        /// </returns>
        public bool IsSupportedType(int variantType)
        {
            for (int i = 0; i < SUPPORTED_TYPES.Length; i++)
                if (variantType == SUPPORTED_TYPES[i])
                    return true;
            return false;
        }



        /// <summary>
        /// Reads a variant type from a byte array
        /// </summary>
        /// <param name="src">The byte array</param>
        /// <param name="offset">The offset in the byte array where the variant starts</param>
        /// <param name="Length">The Length of the variant including the variant type field</param>
        /// <param name="type">The variant type To Read</param>
        /// <param name="codepage">The codepage To use for non-wide strings</param>
        /// <returns>A Java object that corresponds best To the variant field. For
        /// example, a VT_I4 is returned as a {@link long}, a VT_LPSTR as a
        /// {@link String}.</returns>
        public static Object Read(byte[] src, int offset,
                                  int Length, long type,
                                  int codepage)
        {
            Object value;
            int o1 = offset;
            int l1 = Length - LittleEndianConsts.INT_SIZE;
            long lType = type;

            /* Instead of trying To Read 8-bit characters from a Unicode string,
             * Read 16-bit characters. */
            if (codepage == (int)Constants.CP_UNICODE && type == Variant.VT_LPSTR)
                lType = Variant.VT_LPWSTR;

            switch ((int)lType)
            {
                case Variant.VT_EMPTY:
                    {
                        value = null;
                        break;
                    }
                case Variant.VT_I2:
                    {
                        /*
                         * Read a short. In Java it is represented as an
                         * Integer object.
                         */
                        value = LittleEndian.GetShort(src, o1);
                        break;
                    }
                case Variant.VT_I4:
                    {
                        /*
                         * Read a word. In Java it is represented as an
                         * Integer object.
                         */
                        value = LittleEndian.GetInt(src, o1);
                        break;
                    }
                case Variant.VT_I8:
                    {
                        /*
                         * Read a double word. In Java it is represented as a
                         * long object.
                         */
                        value = LittleEndian.GetLong(src, o1);
                        break;
                    }
                case Variant.VT_R8:
                    {
                        /*
                         * Read an eight-byte double value. In Java it is represented as
                         * a Double object.
                         */
                        value = LittleEndian.GetDouble(src, o1);
                        break;
                    }
                case Variant.VT_FILETIME:
                    {
                        /*
                         * Read a FILETIME object. In Java it is represented
                         * as a Date object.
                         */
                        int low = LittleEndian.GetInt(src, o1);
                        o1 += LittleEndianConsts.INT_SIZE;
                        int high = LittleEndian.GetInt(src, o1);
                        if (low == 0 && high == 0)
                            value = null;
                        else
                            value = Util.FiletimeToDate(high, low);
                        break;
                    }
                case Variant.VT_LPSTR:
                    {
                        /*
                         * Read a byte string. In Java it is represented as a
                         * String object. The 0x00 bytes at the end must be
                         * stripped.
                         */
                        int first = o1 + LittleEndianConsts.INT_SIZE;
                        long last = first + LittleEndian.GetUInt(src, o1) - 1;
                        o1 += LittleEndianConsts.INT_SIZE;
                        while (src[(int)last] == 0 && first <= last)
                            last--;
                        int l = (int)(last - first + 1);
                        value = codepage != -1 ?
                            Encoding.GetEncoding(codepage).GetString(src, first, l) :
                            Encoding.UTF8.GetString(src, first, l);
                        break;
                    }
                case Variant.VT_LPWSTR:
                    {
                        /*
                         * Read a Unicode string. In Java it is represented as
                         * a String object. The 0x00 bytes at the end must be
                         * stripped.
                         */
                        int first = o1 + LittleEndianConsts.INT_SIZE;
                        long last = first + LittleEndian.GetUInt(src, o1) - 1;
                        long l = last - first;
                        o1 += LittleEndianConsts.INT_SIZE;
                        StringBuilder b = new StringBuilder((int)(last - first));
                        for (int i = 0; i <= l; i++)
                        {
                            int i1 = o1 + (i * 2);
                            int i2 = i1 + 1;
                            int high = src[i2] << 8;
                            int low = src[i1] & 0x00ff;
                            char c = (char)(high | low);
                            b.Append(c);
                        }
                        /* Strip 0x00 characters from the end of the string: */
                        while (b.Length > 0 && b[b.Length - 1] == 0x00)
                            b.Length = b.Length - 1;
                        value = b.ToString();
                        break;
                    }
                case Variant.VT_CF:
                    {
                        if (l1 < 0)
                        {
                            /**
                             *  YK: reading the ClipboardData packet (VT_CF) is not quite correct.
                             *  The size of the data is determined by the first four bytes of the packet
                             *  while the current implementation calculates it in the Section constructor.
                             *  Test files in Bugzilla 42726 and 45583 clearly show that this approach does not always work.
                             *  The workaround below attempts to gracefully handle such cases instead of throwing exceptions.
                             *
                             *  August 20, 2009
                             */
                            l1 = LittleEndian.GetInt(src, o1); 
                            o1 += LittleEndian.INT_SIZE;
                        }
                        byte[] v = new byte[l1];
                        for (int i = 0; i < l1; i++)
                            v[i] = src[(o1 + i)];
                        value = v;
                        break;
                    }
                case Variant.VT_BOOL:
                    {
                        /*
                         * The first four bytes in src, from src[offset] To
                         * src[offset + 3] contain the DWord for VT_BOOL, so
                         * skip it, we don't need it.
                         */
                        // int first = offset + LittleEndianConstants.INT_SIZE;
                        long boolean = LittleEndian.GetUInt(src, o1);
                        if (boolean != 0)
                            value = true;
                        else
                            value = false;
                        break;
                    }
                default:
                    {
                        byte[] v = new byte[l1];
                        for (int i = 0; i < l1; i++)
                            v[i] = src[(o1 + i)];
                        throw new ReadingNotSupportedException(type, v);
                    }
            }
            return value;
        }


        /// <summary>
        /// Writes a variant value To an output stream. This method ensures that
        /// always a multiple of 4 bytes is written.
        /// If the codepage is UTF-16, which is encouraged, strings
        /// <strong>must</strong> always be written as {@link Variant#VT_LPWSTR}
        /// strings, not as {@link Variant#VT_LPSTR} strings. This method ensure this
        /// by Converting strings appropriately, if needed.
        /// </summary>
        /// <param name="out1">The stream To Write the value To.</param>
        /// <param name="type">The variant's type.</param>
        /// <param name="value">The variant's value.</param>
        /// <param name="codepage">The codepage To use To Write non-wide strings</param>
        /// <returns>The number of entities that have been written. In many cases an
        /// "entity" is a byte but this is not always the case.</returns>
        public static int Write(Stream out1, long type,
                                Object value, int codepage)
        {
            int length = 0;
            switch ((int)type)
            {
                case Variant.VT_BOOL:
                    {
                        int trueOrFalse;
                        if ((bool)value)
                            trueOrFalse = 1;
                        else
                            trueOrFalse = 0;
                        length = TypeWriter.WriteUIntToStream(out1, (uint)trueOrFalse);
                        break;
                    }
                case Variant.VT_LPSTR:
                    {
                        if (codepage == 0)
                            throw new ArgumentOutOfRangeException("codepage");
                        byte[] bytes =
                            (codepage == -1 ?
                            Encoding.UTF8.GetBytes((string)value) :
                            Encoding.GetEncoding(codepage).GetBytes((string)value));
                            
                        length = TypeWriter.WriteUIntToStream(out1, (uint)bytes.Length + 1);
                        byte[] b = new byte[bytes.Length + 1];
                        Array.Copy(bytes, 0, b, 0, bytes.Length);
                        b[b.Length - 1] = 0x00;
                        out1.Write(b, 0, b.Length);
                        length += b.Length;
                        break;
                    }
                case Variant.VT_LPWSTR:
                    {
                        int nrOfChars = ((String)value).Length + 1;
                        length += TypeWriter.WriteUIntToStream(out1, (uint)nrOfChars);
                        char[] s = Util.Pad4((String)value);
                        for (int i = 0; i < s.Length; i++)
                        {
                            int high = ((s[i] & 0x0000ff00) >> 8);
                            int low = (s[i] & 0x000000ff);
                            byte highb = (byte)high;
                            byte lowb = (byte)low;
                            out1.WriteByte(lowb);
                            out1.WriteByte(highb);
                            length += 2;
                        }
                        out1.WriteByte(0x00);
                        out1.WriteByte(0x00);
                        length += 2;
                        break;
                    }
                case Variant.VT_CF:
                    {
                        byte[] b = (byte[])value;
                        out1.Write(b, 0, b.Length);
                        length = b.Length;
                        break;
                    }
                case Variant.VT_EMPTY:
                    {
                        TypeWriter.WriteUIntToStream(out1, Variant.VT_EMPTY);
                        length = LittleEndianConsts.INT_SIZE;
                        break;
                    }
                case Variant.VT_I2:
                    {
                        short x;
                        try
                        {
                            x = Convert.ToInt16(value, CultureInfo.InvariantCulture);
                        }catch(OverflowException)
                        {
                            x=(short)((int)value);
                        }
                        TypeWriter.WriteToStream(out1, x);
                        length = LittleEndianConsts.SHORT_SIZE;
                        break;
                    }
                case Variant.VT_I4:
                    {
                        if (!(value is int))
                        {
                            throw new Exception("Could not cast an object To "
                                    + "int" + ": "
                                    + value.GetType().Name + ", "
                                    + value.ToString());
                        }
                        length += TypeWriter.WriteToStream(out1, (int)value);
                        break;
                    }
                case Variant.VT_I8:
                    {
                        TypeWriter.WriteToStream(out1, Convert.ToInt64(value, CultureInfo.CurrentCulture));
                        length = LittleEndianConsts.LONG_SIZE;
                        break;
                    }
                case Variant.VT_R8:
                    {
                        length += TypeWriter.WriteToStream(out1,
                                  (Double)value);
                        break;
                    }
                case Variant.VT_FILETIME:
                    {
                        long filetime;
                        if (value != null)
                        {
                            filetime = Util.DateToFileTime((DateTime)value);
                        }
                        else
                        {
                            filetime = 0;
                        }
                        int high = (int)((filetime >> 32) & 0x00000000FFFFFFFFL);
                        int low = (int)(filetime & 0x00000000FFFFFFFFL);
                        length += TypeWriter.WriteUIntToStream
                            (out1, (uint)(0x0000000FFFFFFFFL & low));
                        length += TypeWriter.WriteUIntToStream
                            (out1, (uint)(0x0000000FFFFFFFFL & high));
                        
                        break;
                    }
                default:
                    {
                        /* The variant type is not supported yet. However, if the value
                         * is a byte array we can Write it nevertheless. */
                        if (value is byte[])
                        {
                            byte[] b = (byte[])value;
                            out1.Write(b, 0, b.Length);
                            length = b.Length;
                            WriteUnsupportedTypeMessage
                                (new WritingNotSupportedException(type, value));
                        }
                        else
                            throw new WritingNotSupportedException(type, value);
                        break;
                    }
            }

            return length;
        }

    }
}
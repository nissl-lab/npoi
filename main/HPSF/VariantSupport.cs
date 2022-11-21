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

using System.Collections.Generic;

namespace NPOI.HPSF
{
    using System;
    using System.IO;
    using System.Globalization;
    using NPOI.Util;

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
        protected static List<long> unsupportedMessage;

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
                    unsupportedMessage = new List<long>();
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
        /// <param name="length">The Length of the variant including the variant type field</param>
        /// <param name="type">The variant type To Read</param>
        /// <param name="codepage">The codepage To use for non-wide strings</param>
        /// <returns>A Java object that corresponds best To the variant field. For
        /// example, a VT_I4 is returned as a {@link long}, a VT_LPSTR as a
        /// {@link String}.</returns>
        public static Object Read(byte[] src, int offset,
                int length, long type, int codepage)
        {
            TypedPropertyValue typedPropertyValue = new TypedPropertyValue(
                    (int)type, null);
            int unpadded;
            try
            {
                unpadded = typedPropertyValue.ReadValue(src, offset);
            }
            catch (InvalidOperationException)
            {
                int propLength = Math.Min(length, src.Length - offset);
                byte[] v = new byte[propLength];
                System.Array.Copy(src, offset, v, 0, propLength);
                throw new ReadingNotSupportedException(type, v);
            }

            switch ((int)type)
            {
                case Variant.VT_EMPTY:
                case Variant.VT_I4:
                case Variant.VT_I8:
                case Variant.VT_R8:
                    /*
                     * we have more property types that can be converted into Java
                     * objects, but current API need to be preserved, and it returns
                     * other types as byte arrays. In future major versions it shall be
                     * changed -- sergey
                     */
                    return typedPropertyValue.Value;

                case Variant.VT_I2:
                    {
                        /*
                         * also for backward-compatibility with prev. versions of POI
                         * --sergey
                         */
                        return (short)typedPropertyValue.Value;
                    }
                case Variant.VT_FILETIME:
                    {
                        Filetime filetime = (Filetime)typedPropertyValue.Value;
                        return Util.FiletimeToDate((int)filetime.High,
                                (int)filetime.Low);
                    }
                case Variant.VT_LPSTR:
                    {
                        CodePageString string1 = (CodePageString)typedPropertyValue.Value;
                        return string1.GetJavaValue(codepage);
                    }
                case Variant.VT_LPWSTR:
                    {
                        UnicodeString string1 = (UnicodeString)typedPropertyValue.Value;
                        return string1.ToJavaString();
                    }
                case Variant.VT_CF:
                    {
                        // if(l1 < 0) {
                        /*
                         * YK: reading the ClipboardData packet (VT_CF) is not quite
                         * correct. The size of the data is determined by the first four
                         * bytes of the packet while the current implementation calculates
                         * it in the Section constructor. Test files in Bugzilla 42726 and
                         * 45583 clearly show that this approach does not always work. The
                         * workaround below attempts to gracefully handle such cases instead
                         * of throwing exceptions.
                         * 
                         * August 20, 2009
                         */
                        // l1 = LittleEndian.getInt(src, o1); o1 += LittleEndian.INT_SIZE;
                        // }
                        // final byte[] v = new byte[l1];
                        // System.arraycopy(src, o1, v, 0, v.length);
                        // value = v;
                        // break;
                        ClipboardData clipboardData = (ClipboardData)typedPropertyValue.Value;
                        return clipboardData.ToByteArray();
                    }

                case Variant.VT_BOOL:
                    {
                        VariantBool bool1 = (VariantBool)typedPropertyValue.Value;
                        return (bool)bool1.Value;
                    }

                default:
                    {
                        /*
                         * it is not very good, but what can do without breaking current
                         * API? --sergey
                         */
                        byte[] v = new byte[unpadded];
                        System.Array.Copy(src, offset, v, 0, unpadded);
                        throw new ReadingNotSupportedException(type, v);
                    }
            }
        }

        /**
         * <p>Turns a codepage number into the equivalent character encoding's
         * name.</p>
         *
         * @param codepage The codepage number
         *
         * @return The character encoding's name. If the codepage number is 65001,
         * the encoding name is "UTF-8". All other positive numbers are mapped to
         * "cp" followed by the number, e.g. if the codepage number is 1252 the
         * returned character encoding name will be "cp1252".
         *
         * @exception UnsupportedEncodingException if the specified codepage is
         * less than zero.
         */
        public static String CodepageToEncoding(int codepage)
        {
            return CodePageUtil.CodepageToEncoding(codepage);
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
                        byte[] data = new byte[2];
                        if ((bool)value)
                        {
                            out1.WriteByte(0xFF);
                            out1.WriteByte(0xFF);
                        }
                        else
                        {
                            out1.WriteByte(0x00);
                            out1.WriteByte(0x00);
                        }
                        length += 2;
                        break;
                    }
                case Variant.VT_LPSTR:
                    {
                        CodePageString codePageString = new CodePageString((String)value,
                        codepage);
                        length += codePageString.Write(out1);
                        break;
                    }
                case Variant.VT_LPWSTR:
                    {
                        int nrOfChars = ((String)value).Length + 1;
                        length += TypeWriter.WriteUIntToStream(out1, (uint)nrOfChars);
                        char[] s = ((String)value).ToCharArray();
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
                        // NullTerminator
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
                        length += TypeWriter.WriteUIntToStream(out1, Variant.VT_EMPTY);
                        break;
                    }
                case Variant.VT_I2:
                    {
                        short x;
                        try
                        {
                            x = Convert.ToInt16(value, CultureInfo.InvariantCulture);
                        }
                        catch (OverflowException)
                        {
                            x = (short)((int)value);
                        }
                        length += TypeWriter.WriteToStream(out1, x);
                        //length = LittleEndianConsts.SHORT_SIZE;
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
                        length += TypeWriter.WriteToStream(out1, Convert.ToInt64(value, CultureInfo.CurrentCulture));
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
                        Filetime filetimeValue = new Filetime(low, high);
                        length += filetimeValue.Write(out1);
                        //length += TypeWriter.WriteUIntToStream
                        //    (out1, (uint)(0x0000000FFFFFFFFL & low));
                        //length += TypeWriter.WriteUIntToStream
                        //    (out1, (uint)(0x0000000FFFFFFFFL & high));

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
            /* pad values to 4-bytes */
            while ((length & 0x3) != 0)
            {
                out1.WriteByte(0x00);
                length++;
            }
            return length;
        }

    }
}
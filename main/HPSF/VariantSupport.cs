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

namespace NPOI.HPSF
{
    using NPOI.Util;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;

    /// <summary>
    /// <para>
    /// Supports reading and writing of variant data.
    /// </para>
    /// <para>
    /// <strong>FIXME (3):</strong> Reading and writing should be made more
    /// </para>
    /// <para>
    /// uniform than it is now. The following items should be resolved:
    /// </para>
    /// <para>
    /// Reading requires a length parameter that is 4 byte greater than the
    /// actual data, because the variant type field is included.
    /// </para>
    /// <para>
    /// Reading reads from a byte array while writing writes to an byte array
    /// output stream.
    /// </para>
    /// </summary>
    public class VariantSupport : Variant
    {
        /// <summary>
        /// HPSF is able to read these {@link Variant} types.
        /// </summary>
        public static int[] SUPPORTED_TYPES = { Variant.VT_EMPTY,
            Variant.VT_I2, Variant.VT_I4, Variant.VT_I8, Variant.VT_R8,
            Variant.VT_FILETIME, Variant.VT_LPSTR, Variant.VT_LPWSTR,
            Variant.VT_CF, Variant.VT_BOOL };


        //private static POILogger logger = POILogFactory.GetLogger(VariantSupport.class);

        /// <summary>
        /// Keeps a list of the variant types an "unsupported" message has already
        /// been issued for.
        /// </summary>
        private static List<long> unsupportedMessage;

        /// <summary>
        /// Checks whether logging of unsupported variant types warning is turned
        /// on or off.
        /// </summary>
        /// <return>true} if logging is turned on, else
        /// <c>false</c>.
        /// </return>
        public static bool IsLogUnsupportedTypes { get; set; }


        /// <summary>
        /// Writes a warning to {@code System.err} that a variant type is
        /// unsupported by HPSF. Such a warning is written only once for each variant
        /// type. Log messages can be turned on or off by
        /// </summary>
        /// <param name="ex">The exception to log</param>
        protected internal static void WriteUnsupportedTypeMessage
            (UnsupportedVariantTypeException ex)
        {
            if(IsLogUnsupportedTypes)
            {
                if(unsupportedMessage == null)
                    unsupportedMessage = new List<long>();
                long vt = ex.VariantType;
                if(!unsupportedMessage.Contains(vt))
                {
                    //logger.log(POILogger.ERROR, ex.GetMessage());
                    unsupportedMessage.Add(vt);
                }
            }
        }



        /// <summary>
        /// Checks whether HPSF supports the specified variant type. Unsupported
        /// types should be implemented included in the <see cref="SUPPORTED_TYPES"/>
        /// array.
        /// </summary>
        /// @see Variant
        /// <param name="variantType">the variant type to check</param>
        /// <return>true if HPFS supports this type, else
        /// <c>false</c>
        /// </return>
        public bool IsSupportedType(int variantType)
        {
            for(int i = 0; i < SUPPORTED_TYPES.Length; i++)
                if(variantType == SUPPORTED_TYPES[i])
                    return true;
            return false;
        }



        /// <summary>
        /// Reads a variant type from a byte array.
        /// </summary>
        /// <param name="src">The byte array</param>
        /// <param name="offset">The offset in the byte array where the variant starts</param>
        /// <param name="length">The length of the variant including the variant type field</param>
        /// <param name="type">The variant type to read</param>
        /// <param name="codepage">The codepage to use for non-wide strings</param>
        /// <return>Java object that corresponds best to the variant field. For
        /// example, a VT_I4 is returned as a {@link Long}, a VT_LPSTR as a
        /// {@link String}.
        /// </return>
        /// <exception cref="ReadingNotSupportedException">if a property is to be written
        /// who's variant type HPSF does not yet support
        /// </exception>
        /// <exception cref="UnsupportedEncodingException">if the specified codepage is not
        /// supported.
        /// </exception>
        /// @see Variant
        public static Object Read(byte[] src, int offset,
                 int length, long type, int codepage)
        {

            TypedPropertyValue typedPropertyValue = new TypedPropertyValue((int)type, null);
            int unpadded;
            try
            {
                unpadded = typedPropertyValue.ReadValue(src, offset);
            }
            catch(InvalidOperationException)
            {
                int propLength = Math.Min(length, src.Length - offset);
                byte[] v = new byte[propLength];
                System.Array.Copy(src, offset, v, 0, propLength);
                throw new ReadingNotSupportedException(type, v);
            }

            switch((int) type)
            {
                /*
                 * we have more property types that can be converted into Java
                 * objects, but current API need to be preserved, and it returns
                 * other types as byte arrays. In future major versions it shall be
                 * changed -- sergey
                 */
                case Variant.VT_EMPTY:
                case Variant.VT_I4:
                case Variant.VT_I8:
                case Variant.VT_R8:
                    return typedPropertyValue.Value;

                /*
                 * also for backward-compatibility with prev. versions of POI
                 * --sergey
                 */
                case Variant.VT_I2:
                    return ((short) typedPropertyValue.Value);

                case Variant.VT_FILETIME:
                    Filetime filetime = (Filetime)typedPropertyValue.Value;
                    return Util.FiletimeToDate((int) filetime.High, (int) filetime.Low);

                case Variant.VT_LPSTR:
                    CodePageString cpString = (CodePageString)typedPropertyValue.Value;
                    return cpString.GetJavaValue(codepage);

                case Variant.VT_LPWSTR:
                    UnicodeString uniString = (UnicodeString)typedPropertyValue.Value;
                    return uniString.ToJavaString();

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
                // l1 = LittleEndian.GetInt(src, o1); o1 += LittleEndian.INT_SIZE;
                // }
                // final byte[] v = new byte[l1];
                // Array.Copy(src, o1, v, 0, v.Length);
                // value = v;
                // break;
                case Variant.VT_CF:
                    ClipboardData clipboardData = (ClipboardData)typedPropertyValue.Value;
                    return clipboardData.ToByteArray();

                case Variant.VT_BOOL:
                    VariantBool bool1 = (VariantBool)typedPropertyValue.Value;
                    return bool1.Value;

                /*
                 * it is not very good, but what can do without breaking current
                 * API? --sergey
                 */
                default:
                    byte[] v = new byte[unpadded];
                    System.Array.Copy(src, offset, v, 0, unpadded);
                    throw new ReadingNotSupportedException(type, v);
            }
        }

        /// <summary>
        /// Turns a codepage number into the equivalent character encoding's
        /// name.
        /// </summary>
        /// <param name="codepage">The codepage number</param>
        /// 
        /// <return>character encoding's name. If the codepage number is 65001,
        /// the encoding name is "UTF-8". All other positive numbers are mapped to
        /// "cp" followed by the number, e.g. if the codepage number is 1252 the
        /// returned character encoding name will be "cp1252".
        /// </return>

        /// <exception cref="UnsupportedEncodingException">if the specified codepage is
        /// less than zero.
        /// </exception>
        // @Removal(version="3.18")
        [Obsolete("POI 3.16 - use CodePageUtil.CodepageToEncoding(int)")]
        public static String CodepageToEncoding(int codepage)
        {
            return CodePageUtil.CodepageToEncoding(codepage);
        }


        /// <summary>
        /// <para>
        /// Writes a variant value to an output stream. This method ensures that
        /// </para>
        /// <para>
        /// always a multiple of 4 bytes is written.
        /// </para>
        /// <para>
        /// If the codepage is UTF-16, which is encouraged, strings
        /// <strong>must</strong> always be written as {@link Variant#VT_LPWSTR}
        /// strings, not as {@link Variant#VT_LPSTR} strings. This method ensure this
        /// by converting strings appropriately, if needed.
        /// </para>
        /// </summary>
        /// <param name="out">The stream to write the value to.</param>
        /// <param name="type">The variant's type.</param>
        /// <param name="value">The variant's value.</param>
        /// <param name="codepage">The codepage to use to write non-wide strings</param>
        /// <return>number of entities that have been written. In many cases an
        /// "entity" is a byte but this is not always the case.
        /// </return>
        /// <exception cref="IOException">if an I/O exceptions occurs</exception>
        /// <exception cref="WritingNotSupportedException">if a property is to be written
        /// who's variant type HPSF does not yet support
        /// </exception>
        public static int Write(Stream out1, long type,
                                 Object value, int codepage)
        {

            int length = 0;
            switch((int) type)
            {
                case Variant.VT_BOOL:
                    if((bool) value)
                    {
                        out1.WriteByte(0xff);
                        out1.WriteByte(0xff);
                    }
                    else
                    {
                        out1.WriteByte(0x00);
                        out1.WriteByte(0x00);
                    }
                    length += 2;
                    break;

                case Variant.VT_LPSTR:
                    CodePageString codePageString = new CodePageString((String)value, codepage);
                    length += codePageString.Write(out1);
                    break;

                case Variant.VT_LPWSTR:
                    int nrOfChars = ((String)value).Length + 1;
                    length += TypeWriter.WriteUIntToStream(out1, (uint) nrOfChars);
                    foreach(char s in ((String) value).ToCharArray())
                    {
                        int high1 = ((s & 0x0000ff00) >> 8);
                        int low1 = (s & 0x000000ff);
                        byte highb = (byte)high1;
                        byte lowb = (byte)low1;
                        out1.WriteByte(lowb);
                        out1.WriteByte(highb);
                        length += 2;
                    }
                    // NullTerminator
                    out1.WriteByte(0x00);
                    out1.WriteByte(0x00);
                    length += 2;
                    break;

                case Variant.VT_CF:
                    byte[] cf = (byte[])value;
                    out1.Write(cf, 0, cf.Length);
                    length = cf.Length;
                    break;

                case Variant.VT_EMPTY:
                    length += TypeWriter.WriteUIntToStream(out1, Variant.VT_EMPTY);
                    break;

                case Variant.VT_I2:
                    short x;
                    try
                    {
                        x = Convert.ToInt16(value, CultureInfo.InvariantCulture);
                    }
                    catch(OverflowException)
                    {
                        x = (short) ((int) value);
                    }
                    length += TypeWriter.WriteToStream(out1, x);
                    break;

                case Variant.VT_I4:
                    if(value is not int)
                    {
                        throw new InvalidCastException("Could not cast an object to "
                                + typeof(Int32).ToString() + ": "
                                + value.GetType().ToString() + ", "
                                + value.ToString());
                    }
                    length += TypeWriter.WriteToStream(out1, Convert.ToInt32(value));
                    break;

                case Variant.VT_I8:
                    length += TypeWriter.WriteToStream(out1, Convert.ToInt64(value));
                    break;

                case Variant.VT_R8:
                    length += TypeWriter.WriteToStream(out1, Convert.ToDouble(value));
                    break;

                case Variant.VT_FILETIME:
                    long filetime = Util.DateToFileTime(Convert.ToDateTime(value));
                    int high = (int)((filetime >> 32) & 0x00000000FFFFFFFFL);
                    int low = (int)(filetime & 0x00000000FFFFFFFFL);
                    Filetime filetimeValue = new Filetime(low, high);
                    length += filetimeValue.Write(out1);
                    break;

                default:
                    /* The variant type is not supported yet. However, if the value
                     * is a byte array we can write it nevertheless. */
                    if(value is byte[] bytes)
                    {
                        out1.Write(bytes, 0, bytes.Length);
                        length = bytes.Length;
                        WriteUnsupportedTypeMessage(new WritingNotSupportedException(type, bytes));
                    }
                    else
                    {
                        throw new WritingNotSupportedException(type, value);
                    }
                    break;
            }

            /* pad values to 4-bytes */
            while((length & 0x3) != 0)
            {
                out1.WriteByte(0x00);
                length++;
            }

            return length;
        }
    }
}


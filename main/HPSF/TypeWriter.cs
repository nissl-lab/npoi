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
    using System.IO;
    using NPOI.Util;

    /// <summary>
    /// Class for writing little-endian data and more.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2003-02-20 
    /// </summary>
    public class TypeWriter
    {

        /// <summary>
        /// Writes a two-byte value (short) To an output stream.
        /// </summary>
        /// <param name="out1">The stream To Write To..</param>
        /// <param name="n">The number of bytes that have been written.</param>
        /// <returns></returns>
        public static int WriteToStream(Stream out1, short n)
        {
            LittleEndian.PutShort( out1, n ); // FIXME: unsigned
            return LittleEndian.SHORT_SIZE;
        }



        /**
         * Writes a four-byte value To an output stream.
         *
         * @param out The stream To Write To.
         * @param n The value To Write.
         * @exception IOException if an I/O error occurs
         * @return The number of bytes written To the output stream. 
         */
        public static int WriteToStream(Stream out1, int n)
        {
            LittleEndian.PutInt( n, out1 );
            return LittleEndian.INT_SIZE;
        }
        /**
         * Writes a four-byte value To an output stream.
         *
         * @param out The stream To Write To.
         * @param n The value To Write.
         * @exception IOException if an I/O error occurs
         * @return The number of bytes written To the output stream. 
         */
        [Obsolete]
        public static int WriteToStream(Stream out1, uint n)
        {
            int l = LittleEndianConsts.INT_SIZE;
            byte[] buffer = new byte[l];
            LittleEndian.PutUInt(buffer, 0, n);
            out1.Write(buffer, 0, l);
            return l;

        }


        /**
         * Writes a eight-byte value To an output stream.
         *
         * @param out The stream To Write To.
         * @param n The value To Write.
         * @exception IOException if an I/O error occurs
         * @return The number of bytes written To the output stream. 
         */
        public static int WriteToStream(Stream out1, long n)
        {
            LittleEndian.PutLong( n, out1 );
            return LittleEndian.LONG_SIZE;
        }



        /**
         * Writes an unsigned two-byte value To an output stream.
         *
         * @param out The stream To Write To
         * @param n The value To Write
         * @exception IOException if an I/O error occurs
         */
        public static void WriteUShortToStream(Stream out1, int n)
        {
            int high = n & unchecked((int)0xFFFF0000);
            if (high != 0)
                throw new IllegalPropertySetDataException
                    ("Value " + n + " cannot be represented by 2 bytes.");
            //WriteToStream(out1, (short)n);
            LittleEndian.PutUShort( n, out1 );
        }



        /**
         * Writes an unsigned four-byte value To an output stream.
         *
         * @param out The stream To Write To.
         * @param n The value To Write.
         * @return The number of bytes that have been written To the output stream.
         * @exception IOException if an I/O error occurs
         */
        public static int WriteUIntToStream(Stream out1, uint n)
        {
            ulong high = (ulong)(n & unchecked((long)0xFFFFFFFF00000000L));
            if (high != 0 && high != 0xFFFFFFFF00000000L)
                throw new IllegalPropertySetDataException
                    ("Value " + n + " cannot be represented by 4 bytes.");
            LittleEndian.PutUInt( n, out1 );
            return LittleEndian.INT_SIZE;
        }



        /**
         * Writes a 16-byte {@link ClassID} To an output stream.
         *
         * @param out The stream To Write To
         * @param n The value To Write
         * @return The number of bytes written
         * @exception IOException if an I/O error occurs
         */
        public static int WriteToStream(Stream out1, ClassID n)
        {
            byte[] b = new byte[16];
            n.Write(b, 0);
            out1.Write(b, 0, b.Length);
            return b.Length;
        }



        /**
         * Writes an array of {@link Property} instances To an output stream
         * according To the Horrible Property  Format.
         * 
         * @param out The stream To Write To
         * @param properties The array To Write To the stream
         * @param codepage The codepage number To use for writing strings
         * @exception IOException if an I/O error occurs
         * @throws UnsupportedVariantTypeException if HPSF does not support some
         *         variant type.
         */
        public static void WriteToStream(Stream out1,
                                         Property[] properties,
                                         int codepage)
        {
            /* If there are no properties don't Write anything. */
            if (properties == null)
                return;

            /* Write the property list. This is a list containing pairs of property
             * ID and offset into the stream. */
            for (int i = 0; i < properties.Length; i++)
            {
                Property p = properties[i];
                WriteUIntToStream(out1, (uint)p.ID);
                WriteUIntToStream(out1, (uint)p.Count);
            }

            /* Write the properties themselves. */
            for (int i = 0; i < properties.Length; i++)
            {
                Property p = properties[i];
                long type = p.Type;
                WriteUIntToStream(out1, (uint)type);
                VariantSupport.Write(out1, (int)type, p.Value, codepage);
            }
        }



        /**
         * Writes a double value value To an output stream.
         *
         * @param out The stream To Write To.
         * @param n The value To Write.
         * @exception IOException if an I/O error occurs
         * @return The number of bytes written To the output stream. 
         */
        public static int WriteToStream(Stream out1, double n)
        {
            LittleEndian.PutDouble( n, out1 );
            return LittleEndian.DOUBLE_SIZE;
        }

    }
}
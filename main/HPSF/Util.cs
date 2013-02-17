/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding Copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a Copy of the License at

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


    /// <summary>
    /// Provides various static utility methods.
    /// @author Rainer Klute (klute@rainer-klute.de)
    /// @since 2002-02-09
    /// </summary>
    public class Util
    {


        /// <summary>
        /// Copies a part of a byte array into another byte array.
        /// </summary>
        /// <param name="src">The source byte array.</param>
        /// <param name="srcOffSet">OffSet in the source byte array.</param>
        /// <param name="Length">The number of bytes To Copy.</param>
        /// <param name="dst">The destination byte array.</param>
        /// <param name="dstOffSet">OffSet in the destination byte array.</param>
        public static void Copy(byte[] src, int srcOffSet,
                                int Length, byte[] dst,
                                int dstOffSet)
        {
            for (int i = 0; i < Length; i++)
                dst[dstOffSet + i] = src[srcOffSet + i];
        }



        /// <summary>
        /// Concatenates the contents of several byte arrays into a
        /// single one.
        /// </summary>
        /// <param name="byteArrays">The byte arrays To be conCatened.</param>
        /// <returns>A new byte array containing the conCatenated byte arrays.</returns>
        public static byte[] Cat(byte[][] byteArrays)
        {
            int capacity = 0;
            for (int i = 0; i < byteArrays.Length; i++)
                capacity += byteArrays[i].Length;
            byte[] result = new byte[capacity];
            int r = 0;
            for (int i = 0; i < byteArrays.Length; i++)
                for (int j = 0; j < byteArrays[i].Length; j++)
                    result[r++] = byteArrays[i][j];
            return result;
        }



        /// <summary>
        /// Copies bytes from a source byte array into a new byte
        /// array.
        /// </summary>
        /// <param name="src">Copy from this byte array.</param>
        /// <param name="offset">Start Copying here.</param>
        /// <param name="Length">Copy this many bytes.</param>
        /// <returns>The new byte array. Its Length is number of copied bytes.</returns>
        public static byte[] Copy(byte[] src, int offset,
                                  int Length)
        {
            byte[] result = new byte[Length];
            Copy(src, offset, Length, result, 0);
            return result;
        }



        /**
         * The difference between the Windows epoch (1601-01-01
         * 00:00:00) and the Unix epoch (1970-01-01 00:00:00) in
         * milliseconds: 11644473600000L. (Use your favorite spReadsheet
         * program To verify the correctness of this value. By the way,
         * did you notice that you can tell from the epochs which
         * operating system is the modern one? :-))
         */
        public static readonly long EPOCH_DIFF = new DateTime(1970, 1, 1).Ticks;   //11644473600000L;


        /// <summary>
        /// Converts a Windows FILETIME into a {@link DateTime}. The Windows
        /// FILETIME structure holds a DateTime and time associated with a
        /// file. The structure identifies a 64-bit integer specifying the
        /// number of 100-nanosecond intervals which have passed since
        /// January 1, 1601. This 64-bit value is split into the two double
        /// words stored in the structure.
        /// </summary>
        /// <param name="high">The higher double word of the FILETIME structure.</param>
        /// <param name="low">The lower double word of the FILETIME structure.</param>
        /// <returns>The Windows FILETIME as a {@link DateTime}.</returns>
        public static DateTime FiletimeToDate(int high, int low)
        {
            long filetime = ((long)high << 32) | ((long)low & 0xffffffffL);
            return FiletimeToDate(filetime);
        }

        /// <summary>
        /// Converts a Windows FILETIME into a {@link DateTime}. The Windows
        /// FILETIME structure holds a DateTime and time associated with a
        /// file. The structure identifies a 64-bit integer specifying the
        /// number of 100-nanosecond intervals which have passed since
        /// January 1, 1601.
        /// </summary>
        /// <param name="filetime">The filetime To Convert.</param>
        /// <returns>The Windows FILETIME as a {@link DateTime}.</returns>
        public static DateTime FiletimeToDate(long filetime)
        {
            return DateTime.FromFileTime(filetime);
            //long ms_since_16010101 = filetime / (1000 * 10);
            //long ms_since_19700101 = ms_since_16010101 -EPOCH_DIFF;
            //return new DateTime(ms_since_19700101);
        }

        /// <summary>
        /// Converts a {@link DateTime} into a filetime.
        /// </summary>
        /// <param name="dateTime">The DateTime To be Converted</param>
        /// <returns>The filetime</returns>
        public static long DateToFileTime(DateTime dateTime)
        {
            //long ms_since_19700101 = DateTime.Ticks;
            //long ms_since_16010101 = ms_since_19700101 + EPOCH_DIFF;
            //return ms_since_16010101 * (1000 * 10);
            return dateTime.ToFileTime();
        }

        /// <summary>
        /// Compares To object arrays with regarding the objects' order. For
        /// example, [1, 2, 3] and [2, 1, 3] are equal.
        /// </summary>
        /// <param name="c1">The first object array.</param>
        /// <param name="c2">The second object array.</param>
        /// <returns><c>true</c>
        ///  if the object arrays are equal,
        /// <c>false</c>
        ///  if they are not.</returns>
        public static bool AreEqual(IList c1, IList c2)
        {
            return internalEquals(c1, c2);
        }

        /// <summary>
        /// Internals the equals.
        /// </summary>
        /// <param name="c1">The c1.</param>
        /// <param name="c2">The c2.</param>
        /// <returns></returns>
        private static bool internalEquals(IList c1, IList c2)
        {
            IEnumerator o1 = c1.GetEnumerator();
            while( o1.MoveNext())
            {
                Object obj1 = o1.Current;
                bool matchFound = false;
                IEnumerator o2 = c2.GetEnumerator();
                while( !matchFound && o2.MoveNext())
                {
                    Object obj2 = o2.Current;
                    if (obj1.Equals(obj2))
                    {
                        matchFound = true;
                        //o2[i2] = null;
                    }
                }
                if (!matchFound)
                    return false;
            }
            return true;
        }



        /// <summary>
        /// Pads a byte array with 0x00 bytes so that its Length is a multiple of
        /// 4.
        /// </summary>
        /// <param name="ba">The byte array To pad.</param>
        /// <returns>The padded byte array.</returns>
        public static byte[] Pad4(byte[] ba)
        {
            int PAD = 4;
            byte[] result;
            int l = ba.Length % PAD;
            if (l == 0)
                result = ba;
            else
            {
                l = PAD - l;
                result = new byte[ba.Length+l ];
                System.Array.Copy(ba,result, ba.Length);
            }
            return result;
        }



        /// <summary>
        /// Pads a character array with 0x0000 characters so that its Length is a
        /// multiple of 4.
        /// </summary>
        /// <param name="ca">The character array To pad.</param>
        /// <returns>The padded character array.</returns>
        public static char[] Pad4(char[] ca)
        {
            int PAD = 4;
            char[] result;
            int l = ca.Length % PAD;
            if (l == 0)
                result = ca;
            else
            {
                l = PAD - l;
                result = new char[ca.Length+l];
                System.Array.Copy(ca, result, ca.Length);
            }
            return result;
        }



        /// <summary>
        /// Pads a string with 0x0000 characters so that its Length is a
        /// multiple of 4.
        /// </summary>
        /// <param name="s">The string To pad.</param>
        /// <returns> The padded string as a character array.</returns>
        public static char[] Pad4(String s)
        {
            return Pad4(s.ToCharArray());
        }
    }
}

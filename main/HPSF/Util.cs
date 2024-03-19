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

using System.Collections;

namespace NPOI.HPSF
{
    using System;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Provides various static utility methods.
    /// </summary>
    public class Util
    {
        /// <summary>
        /// The difference between the Windows epoch (1601-01-01
        /// 00:00:00) and the Unix epoch (1970-01-01 00:00:00) in
        /// milliseconds: 11644473600000L. (Use your favorite spreadsheet
        /// program to verify the correctness of this value. By the way,
        /// did you notice that you can tell from the epochs which
        /// operating system is the modern one? :-))
        /// </summary>
        public static long EPOCH_DIFF = 11644473600000L;


        /// <summary>
        /// Converts a Windows FILETIME into a {@link Date}. The Windows
        /// FILETIME structure holds a date and time associated with a
        /// file. The structure identifies a 64-bit integer specifying the
        /// number of 100-nanosecond intervals which have passed since
        /// January 1, 1601. This 64-bit value is split into the two double
        /// words stored in the structure.
        /// </summary>
        /// <param name="high">The higher double word of the FILETIME structure.</param>
        /// <param name="low">The lower double word of the FILETIME structure.</param>
        /// <return>Windows FILETIME as a {@link Date}.</return>
        public static DateTime FiletimeToDate(int high, int low)
        {
            long filetime = ((long)high) << 32 | (low & 0xffffffffL);
            return FiletimeToDate(filetime);
        }

        /// <summary>
        /// Converts a Windows FILETIME into a {@link Date}. The Windows
        /// FILETIME structure holds a date and time associated with a
        /// file. The structure identifies a 64-bit integer specifying the
        /// number of 100-nanosecond intervals which have passed since
        /// January 1, 1601.
        /// </summary>
        /// <param name="filetime">The filetime to convert.</param>
        /// <return>Windows FILETIME as a {@link Date}.</return>
        public static DateTime FiletimeToDate(long filetime)
        {
            return DateTime.FromFileTime(filetime);
            //long ms_since_16010101 = filetime / (1000 * 10);
            //long ms_since_19700101 = ms_since_16010101 - EPOCH_DIFF;
            //return new DateTime(ms_since_19700101);
        }



        /// <summary>
        /// Converts a {@link Date} into a filetime.
        /// </summary>
        /// <param name="date">The date to be converted</param>
        /// <return>filetime</return>
        public static long DateToFileTime(DateTime dateTime)
        {
            //long ms_since_19700101 = DateTime.Ticks;
            //long ms_since_16010101 = ms_since_19700101 + EPOCH_DIFF;
            //return ms_since_16010101 * (1000 * 10);
            return dateTime.ToFileTime();
        }



        /// <summary>
        /// Compares to object arrays with regarding the objects' order. For
        /// example, [1, 2, 3] and [2, 1, 3] are equal.
        /// </summary>
        /// <param name="c1">The first object array.</param>
        /// <param name="c2">The second object array.</param>
        /// <return>if the object arrays are equal,
        /// <code>false</code> if they are not.
        /// </return>
        public static bool AreEqual(IList c1, IList c2)
        {
            for(int i1 = 0; i1 < c1.Count; i1++)
            {
                Object obj1 = c1[i1];
                bool matchFound = false;
                for(int i2 = 0; !matchFound && i2 < c1.Count; i2++)
                {
                    Object obj2 = c2[i2];
                    if(obj1.Equals(obj2))
                    {
                        matchFound = true;
                        //c2[i2] = null;
                    }
                }
                if(!matchFound)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Pads a byte array with 0x00 bytes so that its length is a multiple of
        /// 4.
        /// </summary>
        /// <param name="ba">The byte array to pad.</param>
        /// <return>padded byte array.</return>
        public static byte[] Pad4(byte[] ba)
        {
            int PAD = 4;
            byte[] result;
            int l = ba.Length % PAD;
            if(l == 0)
                result = ba;
            else
            {
                l = PAD - l;
                result = new byte[ba.Length + l];
                System.Array.Copy(ba, 0, result, 0, ba.Length);
            }
            return result;
        }


        /// <summary>
        /// Returns a textual representation of a {@link Throwable}, including a
        /// stacktrace.
        /// </summary>
        /// <param name="t">The {@link Throwable}</param>
        /// 
        /// <return>string containing the output of a call to
        /// <code>t.printStacktrace()</code>.
        /// </return>
        public static String ToString(Exception t)
        {
            StringWriter sw = new StringWriter();
            //PrintWriter pw = new PrintWriter(sw);
            //t.printStackTrace(pw);
            //pw.Close();
            try
            {
                sw.Close();
                return sw.ToString();
            }
            catch(IOException e)
            {
                StringBuilder b = new StringBuilder(t.Message);
                b.Append("\n");
                b.Append("Could not create a stacktrace. Reason: ");
                b.Append(e.Message);
                return b.ToString();
            }
        }

    }
}


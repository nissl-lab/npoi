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


using System;
using System.Text;

namespace NPOI.XSSF.Binary
{
    using NPOI;
    using NPOI.Util;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBUtils
    {
        /// <summary>
        /// Reads an XLNullableWideString.
        /// </summary>
        /// <param name="data">data from which to read</param>
        /// <param name="offset">in data from which to start</param>
        /// <param name="sb">buffer to which to write.  You must SetLength(0) before calling!</param>
        /// <return>number of bytes read</return>
        /// <exception cref="XSSFBParseException">if there was an exception during reading</exception>
        public static int ReadXLNullableWideString(byte[] data, int offset, StringBuilder sb)
        {

            long numChars = LittleEndian.GetUInt(data, offset);
            if(numChars < 0)
            {
                throw new XSSFBParseException("too few chars to read");
            }
            else if(numChars == 0xFFFFFFFFL)
            { //this means null value (2.5.166), do not read any bytes!!!
                return 0;
            }
            else if(numChars > 0xFFFFFFFFL)
            {
                throw new XSSFBParseException("too many chars to read");
            }

            int numBytes = 2*(int)numChars;
            offset += 4;
            if(offset+numBytes > data.Length)
            {
                throw new XSSFBParseException("trying to read beyond data length:" +
                 "offset="+offset+", numBytes="+numBytes+", data.Length="+data.Length);
            }
            
            sb.Append(Encoding.Unicode.GetString(data, offset, numBytes));
            numBytes+=4;
            return numBytes;
        }


        /// <summary>
        /// Reads an XLNullableWideString.
        /// </summary>
        /// <param name="data">data from which to read</param>
        /// <param name="offset">in data from which to start</param>
        /// <param name="sb">buffer to which to write.  You must SetLength(0) before calling!</param>
        /// <return>number of bytes read</return>
        /// <exception cref="XSSFBParseException">if there was an exception while trying to read the string</exception>
        public static int ReadXLWideString(byte[] data, int offset, StringBuilder sb)
        {

            long numChars = LittleEndian.GetUInt(data, offset);
            if(numChars < 0)
            {
                throw new XSSFBParseException("too few chars to read");
            }
            else if(numChars > 0xFFFFFFFFL)
            {
                throw new XSSFBParseException("too many chars to read");
            }
            int numBytes = 2*(int)numChars;
            offset += 4;
            if(offset+numBytes > data.Length)
            {
                throw new XSSFBParseException("trying to read beyond data length");
            }
            sb.Append(Encoding.Unicode.GetString(data, offset, numBytes));
            numBytes+=4;
            return numBytes;
        }

        public static int CastToInt(long val)
        {
            if(val < Int32.MaxValue && val > Int32.MinValue)
            {
                return (int) val;
            }
            throw new POIXMLException("val ("+val+") can't be cast to int");
        }

        public static short CastToShort(int val)
        {
            if(val < short.MaxValue && val > short.MinValue)
            {
                return (short) val;
            }
            throw new POIXMLException("val ("+val+") can't be cast to short");

        }

        //TODO: Move to LittleEndian?
        public static int Get24BitInt(byte[] data, int offset)
        {
            int i = offset;
            int b0 = data[i++] & 0xFF;
            int b1 = data[i++] & 0xFF;
            int b2 = data[i] & 0xFF;
            return (b2 << 16) + (b1 << 8) + b0;
        }
    }
}


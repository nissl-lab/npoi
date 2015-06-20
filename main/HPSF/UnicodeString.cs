/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

using System;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class UnicodeString
    {
        //private final static POILogger logger = POILogFactory
        //   .getLogger( UnicodeString.class );

        private byte[] _value;

        public UnicodeString(byte[] data, int offset)
        {
            int length = LittleEndian.GetInt(data, offset);
            int dataOffset = offset + LittleEndian.INT_SIZE;

            if (!validLength(length, data, dataOffset))
            {
                // If the length looks wrong, this might be because the offset is sometimes expected 
                // to be on a 4 byte boundary. Try checking with that if so, rather than blowing up with
                // and  ArrayIndexOutOfBoundsException below
                bool valid = false;
                int past4byte = offset % 4;
                if (past4byte != 0)
                {
                    offset = offset + past4byte;
                    length = LittleEndian.GetInt(data, offset);
                    dataOffset = offset + LittleEndian.INT_SIZE;

                    valid = validLength(length, data, dataOffset);
                }

                if (!valid)
                {
                    throw new IllegalPropertySetDataException(
                            "UnicodeString started at offset #" + offset +
                            " is not NULL-terminated");
                }
            }

            if (length == 0)
            {
                _value = new byte[0];
                return;
            }

            _value = LittleEndian.GetByteArray(data, dataOffset, length * 2);
        }

        /**
         * Checks to see if the specified length seems valid,
         *  given the amount of data available still to read,
         *  and the requirement that the string be NULL-terminated
         */
        bool validLength(int length, byte[] data, int offset)
        {
            if (length == 0)
            {
                return true;
            }

            int endOffset = offset + (length * 2);
            if (endOffset <= data.Length)
            {
                // Data Length is OK, ensure it's null terminated too
                if (data[endOffset - 1] == 0 && data[endOffset - 2] == 0)
                {
                    // Length looks plausible
                    return true;
                }
            }

            // Something's up/invalid with that length for the given data+offset
            return false;
        }

        public int Size
        {
            get { return LittleEndian.INT_SIZE + _value.Length; }
        }

        public byte[] Value
        {
            get { return _value; }
        }

        public String ToJavaString()
        {
            if (_value.Length == 0)
                return null;

            String result = StringUtil.GetFromUnicodeLE(_value, 0,
                    _value.Length >> 1);

            int terminator = result.IndexOf('\0');
            if (terminator == -1)
            {
                //logger.log(
                //        POILogger.WARN,
                //        "String terminator (\\0) for UnicodeString property value not found."
                //                + "Continue without trimming and hope for the best." );
                return result;
            }
            if (terminator != result.Length - 1)
            {
                //logger.log(
                //        POILogger.WARN,
                //        "String terminator (\\0) for UnicodeString property value occured before the end of string. "
                //                + "Trimming and hope for the best." );
            }
            return result.Substring(0, terminator);
        }
    }
}
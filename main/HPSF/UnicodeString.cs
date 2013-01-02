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

            if (length == 0)
            {
                _value = new byte[0];
                return;
            }

            _value = LittleEndian.GetByteArray(data, offset
                    + LittleEndian.INT_SIZE, length * 2);

            if (_value[length * 2 - 1] != 0 || _value[length * 2 - 2] != 0)
                throw new IllegalPropertySetDataException(
                        "UnicodeString started at offset #" + offset
                                + " is not NULL-terminated");
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
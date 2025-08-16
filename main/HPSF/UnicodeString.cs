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
using System.IO;
using NPOI.Util;

namespace NPOI.HPSF
{
    internal class UnicodeString
    {
        //private final static POILogger logger = POILogFactory
        //   .getLogger( UnicodeString.class );
        internal UnicodeString() {}
        private byte[] _value;

        internal void Read(LittleEndianByteArrayInputStream lei)
        {
            int length = lei.ReadInt();
            int unicodeBytes = length*2;
            _value = new byte[unicodeBytes];
        
            // If Length is zero, this field MUST be zero bytes in length. If Length is
            // nonzero, this field MUST be a null-terminated array of 16-bit Unicode characters, followed by
            // zero padding to a multiple of 4 bytes. The string represented by this field SHOULD NOT
            // contain embedded or additional trailing null characters.
        
            if (length == 0)
            {
                return;
            }

             int offset = lei.GetReadIndex();
        
            lei.ReadFully(_value);

            if (_value[unicodeBytes-2] != 0 || _value[unicodeBytes-1] != 0) 
            {
                string msg = "UnicodeString started at offset #" + offset + " is not NULL-terminated";
                throw new IllegalPropertySetDataException(msg);
            }
        
            TypedPropertyValue.SkipPadding(lei);
        }

        internal byte[] Value
        {
            get { return _value; }
        }

        internal String ToJavaString()
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

        internal void SetJavaValue(string string1 ) 
        {
            _value = CodePageUtil.GetBytesInCodePage(string1 + "\0", CodePageUtil.CP_UNICODE);
        }

        internal int Write( Stream out1 ) 
        {
            LittleEndian.PutUInt( _value.Length / 2, out1 );
            out1.Write( _value, 0, _value.Length );
            return LittleEndianConsts.INT_SIZE + _value.Length;
        }
    }
}
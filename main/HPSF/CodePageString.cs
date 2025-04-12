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
using System.Text;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class CodePageString
    {
        //private final static POILogger logger = POILogFactory
        //   .getLogger( CodePageString.class );

        private byte[] _value;

        internal CodePageString() {}
        internal void Read( LittleEndianByteArrayInputStream lei )
        {
            int offset = lei.GetReadIndex();
            int size = lei.ReadInt();
            _value = new byte[size];
            if (size == 0) {
                return;
            }

            // If Size is zero, this field MUST be zero bytes in length. If Size is
            // nonzero and the CodePage property set's CodePage property has the value CP_WINUNICODE
            // (0x04B0), then the value MUST be a null-terminated array of 16-bit Unicode characters,
            // followed by zero padding to a multiple of 4 bytes. If Size is nonzero and the property set's
            // CodePage property has any other value, it MUST be a null-terminated array of 8-bit characters
            // from the code page identified by the CodePage property, followed by zero padding to a
            // multiple of 4 bytes. The string represented by this field MAY contain embedded or additional
            // trailing null characters and an OLEPS implementation MUST be able to handle such strings.        
        
            lei.ReadFully(_value);
            if (_value[size - 1] != 0 ) {
                // TODO Some files, such as TestVisioWithCodepage.vsd, are currently
                // triggering this for values that don't look like codepages
                // See Bug #52258 for details
                //String msg = "CodePageString started at offset #" + offset + " is not NULL-terminated";
                //LOG.log(POILogger.WARN, msg);
            }

            TypedPropertyValue.SkipPadding(lei);
        }

        [Obsolete]
        public CodePageString(String aString, int codepage)
        {
            SetJavaValue(aString, codepage);
        }

        public String GetJavaValue(int codepage)
        {
#if NETSTANDARD2_1 || NET6_0_OR_GREATER || NETSTANDARD2_0
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif 
            int cp = ( codepage == -1 ) ? Property.DEFAULT_CODEPAGE : codepage;
            String result = CodePageUtil.GetStringFromCodePage(_value, cp);

        
            int terminator = result.IndexOf( '\0' );
            if ( terminator == -1 )
            {
                //String msg = 
                //    "String terminator (\\0) for CodePageString property value not found." +
                //    "Continue without trimming and hope for the best.";
                //LOG.log(POILogger.WARN, msg);
                return result;
            }
            if ( terminator != result.Length - 1 )
            {
                //String msg = 
                //    "String terminator (\\0) for CodePageString property value occured before the end of string. "+
                //    "Trimming and hope for the best.";
                //LOG.log(POILogger.WARN, msg );
            }
            return result.Substring(0, terminator);
        }

        public int Size
        {
            get { return LittleEndianConsts.INT_SIZE + _value.Length; }
        }

        public void SetJavaValue(String aString, int codepage)
        {
            String stringNT = aString + "\0";
            if (codepage == -1)
                _value = Encoding.UTF8.GetBytes(stringNT);
            else
                _value = CodePageUtil.GetBytesInCodePage(stringNT, codepage);
                //_value = Encoding.GetEncoding(codepage).GetBytes(aString + "\0");
            int cp = ( codepage == -1 ) ? Property.DEFAULT_CODEPAGE : codepage;
            _value = CodePageUtil.GetBytesInCodePage(aString + "\0", cp);
        }

        public int Write(Stream out1)
        {
            LittleEndian.PutInt(_value.Length, out1);
            out1.Write(_value, 0, _value.Length);
            return LittleEndian.INT_SIZE + _value.Length;
        }
    }
}
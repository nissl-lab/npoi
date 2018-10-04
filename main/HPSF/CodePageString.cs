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

        public CodePageString(byte[] data, int startOffset)
        {
            int offset = startOffset;

            int size = LittleEndian.GetInt(data, offset);
            offset += LittleEndian.INT_SIZE;

            _value = LittleEndian.GetByteArray(data, offset, size);
            if (size != 0 && _value[size - 1] != 0)
            {
                // TODO Some files, such as TestVisioWithCodepage.vsd, are currently
                //  triggering this for values that don't look like codepages
                // See Bug #52258 for details
                Console.WriteLine("CodePageString started at offset #" + offset
                            + " is not NULL-terminated" );
                //            throw new IllegalPropertySetDataException(
                //                    "CodePageString started at offset #" + offset
                //                            + " is not NULL-terminated" );
            }
        }

        public CodePageString(String aString, int codepage)
        {
            SetJavaValue(aString, codepage);
        }

        public String GetJavaValue(int codepage)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            String result;
            if (codepage == -1)
                result = Encoding.UTF8.GetString(_value);
            else
                result = Encoding.GetEncoding(codepage).GetString(_value);
            int terminator = result.IndexOf('\0');
            if (terminator == -1)
            {
                //logger.log(
                //        POILogger.WARN,
                //        "String terminator (\\0) for CodePageString property value not found."
                //                + "Continue without trimming and hope for the best." );
                return result;
            }
            if (terminator != result.Length - 1)
            {
                //logger.log(
                //        POILogger.WARN,
                //        "String terminator (\\0) for CodePageString property value occured before the end of string. "
                //                + "Trimming and hope for the best." );
            }
            return result.Substring(0, terminator);
        }

        public int Size
        {
            get { return LittleEndian.INT_SIZE + _value.Length; }
        }

        public void SetJavaValue(String aString, int codepage)
        {
            if (codepage == -1)
                _value = Encoding.UTF8.GetBytes(aString + "\0");
            else
                _value = Encoding.GetEncoding(codepage).GetBytes(aString + "\0");
        }

        public int Write(Stream out1)
        {
            LittleEndian.PutInt(_value.Length, out1);
            out1.Write(_value, 0, _value.Length);
            return LittleEndian.INT_SIZE + _value.Length;
        }
    }
}
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
    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    internal class XSSFBRichStr
    {

        public static XSSFBRichStr Build(byte[] bytes, int offset)
        {

            byte first = bytes[offset];
            bool dwSizeStrRunExists = (first >> 7 & 1) == 1;//first bit == 1?
            bool phoneticExists = (first >> 6 & 1) == 1;//second bit == 1?
            StringBuilder sb = new StringBuilder();

            int read = XSSFBUtils.ReadXLWideString(bytes, offset+1, sb);
            //TODO: parse phonetic strings.
            return new XSSFBRichStr(sb.ToString(), "");
        }

        private String @string;
        private String phoneticString;

        public XSSFBRichStr(String @string, String phoneticString)
        {
            this.@string = @string;
            this.phoneticString = phoneticString;
        }

        public String String => @string;
    }
}


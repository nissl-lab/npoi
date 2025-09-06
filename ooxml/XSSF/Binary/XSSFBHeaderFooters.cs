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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XSSF.Binary
{
    using NPOI.Util;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    class XSSFBHeaderFooters
    {

        public static XSSFBHeaderFooters parse(byte[] data)
        {
            //parse these at some point.
            //bool diffOddEven = false;
            //bool diffFirst = false;
            //bool scaleWDoc = false;
            //bool alignMargins = false;

            int offset = 2;
            XSSFBHeaderFooters xssfbHeaderFooter = new XSSFBHeaderFooters();
            xssfbHeaderFooter.header = new XSSFBHeaderFooter("header", true);
            xssfbHeaderFooter.footer = new XSSFBHeaderFooter("footer", false);
            xssfbHeaderFooter.headerEven = new XSSFBHeaderFooter("evenHeader", true);
            xssfbHeaderFooter.footerEven = new XSSFBHeaderFooter("evenFooter", false);
            xssfbHeaderFooter.headerFirst = new XSSFBHeaderFooter("firstHeader", true);
            xssfbHeaderFooter.footerFirst = new XSSFBHeaderFooter("firstFooter", false);
            offset += readHeaderFooter(data, offset, xssfbHeaderFooter.header);
            offset += readHeaderFooter(data, offset, xssfbHeaderFooter.footer);
            offset += readHeaderFooter(data, offset, xssfbHeaderFooter.headerEven);
            offset += readHeaderFooter(data, offset, xssfbHeaderFooter.footerEven);
            offset += readHeaderFooter(data, offset, xssfbHeaderFooter.headerFirst);
            readHeaderFooter(data, offset, xssfbHeaderFooter.footerFirst);
            return xssfbHeaderFooter;
        }

        private static int readHeaderFooter(byte[] data, int offset, XSSFBHeaderFooter headerFooter)
        {
            if(offset + 4 >= data.Length)
            {
                return 0;
            }
            StringBuilder sb = new StringBuilder();
            int bytesRead = XSSFBUtils.ReadXLNullableWideString(data, offset, sb);
            headerFooter.RawString = sb.ToString();
            return bytesRead;
        }

        private XSSFBHeaderFooter header;
        private XSSFBHeaderFooter footer;
        private XSSFBHeaderFooter headerEven;
        private XSSFBHeaderFooter footerEven;
        private XSSFBHeaderFooter headerFirst;
        private XSSFBHeaderFooter footerFirst;

        public XSSFBHeaderFooter Header => header;

        public XSSFBHeaderFooter Footer => footer;

        public XSSFBHeaderFooter HeaderEven => headerEven;

        public XSSFBHeaderFooter FooterEven => footerEven;

        public XSSFBHeaderFooter HeaderFirst => headerFirst;

        public XSSFBHeaderFooter FooterFirst => footerFirst;
    }
}


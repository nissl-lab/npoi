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
    using NPOI.XSSF.UserModel.Helpers;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBHeaderFooter
    {

        private static  HeaderFooterHelper HEADER_FOOTER_HELPER = new HeaderFooterHelper();

        private  String headerFooterTypeLabel;
        private  bool isHeader;
        private String rawString;


        public XSSFBHeaderFooter(String headerFooterTypeLabel, bool isHeader)
        {
            this.headerFooterTypeLabel = headerFooterTypeLabel;
            this.isHeader = isHeader;
        }

        public String HeaderFooterTypeLabel => headerFooterTypeLabel;

        public String RawString
        {
            get => rawString;
            set => rawString = value;
        }

        public String String
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                String left = HEADER_FOOTER_HELPER.GetLeftSection(rawString);
                String center = HEADER_FOOTER_HELPER.GetCenterSection(rawString);
                String right = HEADER_FOOTER_HELPER.GetRightSection(rawString);
                if(left != null && left.Length > 0)
                {
                    sb.Append(left);
                }
                if(center != null && center.Length > 0)
                {
                    if(sb.Length > 0)
                    {
                        sb.Append(" ");
                    }
                    sb.Append(center);
                }
                if(right != null && right.Length > 0)
                {
                    if(sb.Length > 0)
                    {
                        sb.Append(" ");
                    }
                    sb.Append(right);
                }
                return sb.ToString();
            }
            set
            {
                this.rawString = value;
            }
        }

        public bool IsHeader => isHeader;

    }
}


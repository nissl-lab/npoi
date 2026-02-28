/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel.Extensions;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
namespace NPOI.XSSF.UserModel
{

    /**
     * 
     * First page footer content. Corresponds to first printed page.  
     * The first logical page in the sheet may not be printed, for example, if the print area is specified to 
     * be a range such that it falls outside the first page's scope.
     * 
     */
    public class XSSFFirstFooter : XSSFHeaderFooter, IFooter
    {

        /**
         * Create an instance of XSSFFirstFooter from the supplied XML bean
         * @see XSSFSheet#getFirstFooter()
         * @param headerFooter
         */
        public XSSFFirstFooter(CT_HeaderFooter headerFooter)
            : base(headerFooter)
        {

            headerFooter.differentFirst = (true);
        }

        /**
         * Get the content text representing the footer
         * @return text
         */
        public override String Text
        {
            get
            {
                return GetHeaderFooter().firstFooter;
            }
            set 
            {
                if (value == null)
                {
                    GetHeaderFooter().firstFooter = null;
                }
                else
                {
                    GetHeaderFooter().firstFooter = (value);
                }
            }
        }
    }
}



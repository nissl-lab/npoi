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

using NPOI.XSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
namespace NPOI.XSSF.UserModel.Helpers
{


    /**
     * 
     * This class is a wrapper around the CT_XmlColumnPr (Open Office XML Part 4:
     * chapter 3.5.1.7)
     * 
     *
     * @author Roberto Manicardi
     */
    public class XSSFXmlColumnPr
    {

        private XSSFTable table;
        private CT_TableColumn ctTableColumn;
        private CT_XmlColumnPr ctXmlColumnPr;

        public XSSFXmlColumnPr(XSSFTable table, CT_TableColumn ctTableColum, CT_XmlColumnPr CT_XmlColumnPr)
        {
            this.table = table;
            this.ctTableColumn = ctTableColum;
            this.ctXmlColumnPr = CT_XmlColumnPr;
        }

        public long GetMapId()
        {
            return ctXmlColumnPr.mapId;
        }

        public String GetXPath()
        {
            return ctXmlColumnPr.xpath;
            
        }
        /**
         * (see Open Office XML Part 4: chapter 3.5.1.3)
         * @return An integer representing the unique identifier of this column. 
         */
        public long GetId()
        {
            return ctTableColumn.id;
        }


        /**
         * If the XPath is, for example, /Node1/Node2/Node3 and /Node1/Node2 is the common XPath for the table, the local XPath is /Node3
         * 	
         * @return the local XPath 
         */
        public String GetLocalXPath()
        {
            String localXPath = "";
            int numberOfCommonXPathAxis = table.GetCommonXpath().Split(new char[] { '/' }).Length - 1;

            String[] xPathTokens = ctXmlColumnPr.xpath.Split(new char[] { '/' });
            for (int i = numberOfCommonXPathAxis; i < xPathTokens.Length; i++)
            {
                localXPath += "/" + xPathTokens[i];
            }
            return localXPath;
        }

        public ST_XmlDataType GetXmlDataType()
        {

            return ctXmlColumnPr.xmlDataType;
        }
    }

}



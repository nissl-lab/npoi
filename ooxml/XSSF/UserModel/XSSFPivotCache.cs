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
namespace NPOI.XSSF.UserModel
{
    using System;
    using NPOI.OpenXml4Net.OPC;
    using System.IO;
    using System.Xml;
    using NPOI.OpenXmlFormats.Spreadsheet;

    public class XSSFPivotCache : POIXMLDocumentPart
    {

        private CT_PivotCache ctPivotCache;


        public XSSFPivotCache()
            : base()
        {
            ctPivotCache = new CT_PivotCache();
        }


        public XSSFPivotCache(CT_PivotCache ctPivotCache)
            : base()
        {
            this.ctPivotCache = ctPivotCache;
        }

        /**
        * Creates n XSSFPivotCache representing the given package part and relationship.
        * Should only be called when Reading in an existing file.
        *
        * @param part - The package part that holds xml data representing this pivot cache defInition.
        * @param rel - the relationship of the given package part in the underlying OPC package
        */

        protected XSSFPivotCache(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            ReadFrom(part.GetInputStream());
        }


        protected void ReadFrom(Stream is1)
        {
            try
            {
                //XmlOptions options = new XmlOptions(DEFAULT_XML_OPTIONS);
                ////Removing root element
                //options.LoadReplaceDocumentElement = (/*setter*/null);
                XmlDocument xmlDoc = ConvertStreamToXml(is1);
                ctPivotCache = CT_PivotCache.Parse(xmlDoc.DocumentElement, NamespaceManager);
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }


        public CT_PivotCache GetCTPivotCache()
        {
            return ctPivotCache;
        }
    }
}
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


    public class XSSFPivotCacheRecords : POIXMLDocumentPart
    {
        private CT_PivotCacheRecords ctPivotCacheRecords;


        public XSSFPivotCacheRecords()
            : base()
        {

            ctPivotCacheRecords = new CT_PivotCacheRecords();
        }

        /**
         * Creates an XSSFPivotCacheRecords representing the given package part and relationship.
         * Should only be called when Reading in an existing file.
         *
         * @param part - The package part that holds xml data representing this pivot cache records.
         * @param rel - the relationship of the given package part in the underlying OPC package
         */

        protected XSSFPivotCacheRecords(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            ReadFrom(part.GetInputStream());
        }


        protected void ReadFrom(Stream is1)
        {
            try
            {
                //XmlOptions options = new XmlOptions(DEFAULT_XML_OPTIONS);
                //Removing root element
                //options.LoadReplaceDocumentElement = (/*setter*/null);
                XmlDocument xmldoc = ConvertStreamToXml(is1);
                ctPivotCacheRecords = CT_PivotCacheRecords.Parse(xmldoc.DocumentElement, NamespaceManager);
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }



        public CT_PivotCacheRecords GetCtPivotCacheRecords()
        {
            return ctPivotCacheRecords;
        }



        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //Sets the pivotCacheDefInition tag
            //xmlOptions.SetSaveSyntheticDocumentElement(new QName(CTPivotCacheRecords.type.Name.
            //        GetNamespaceURI(), "pivotCacheRecords"));
            ctPivotCacheRecords.Save(out1);
            out1.Close();
        }
    }
}
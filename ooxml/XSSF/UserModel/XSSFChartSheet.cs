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

using NPOI.OpenXmlFormats.Spreadsheet;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System;
using NPOI.OpenXmlFormats;
using NPOI.Util;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XSSF.UserModel
{

    /**
     * High level representation of Sheet Parts that are of type 'chartsheet'.
     * <p>
     *  Chart sheet is a special kind of Sheet that Contains only chart and no data.
     * </p>
     *
     * @author Yegor Kozlov
     */
    public class XSSFChartSheet : XSSFSheet
    {

        private static byte[] BLANK_WORKSHEET = blankWorksheet();

        protected CT_Chartsheet chartsheet;

        protected XSSFChartSheet(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

        }

        internal override void Read(Stream is1)
        {
            //Initialize the supeclass with a blank worksheet
            base.Read(new MemoryStream(BLANK_WORKSHEET));

            try
            {
                XmlDocument doc = ConvertStreamToXml(is1);
                chartsheet = ChartsheetDocument.Parse(doc, XSSFSheet.NamespaceManager).GetChartsheet();
            }
            catch (XmlException e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Provide access to the CTChartsheet bean holding this sheet's data
         *
         * @return the CTChartsheet bean holding this sheet's data
         */
        public CT_Chartsheet GetCTChartsheet()
        {
            return chartsheet;
        }


        protected override NPOI.OpenXmlFormats.Spreadsheet.CT_Drawing GetCTDrawing()
        {
            return chartsheet.drawing;
        }


        protected override NPOI.OpenXmlFormats.Spreadsheet.CT_LegacyDrawing GetCTLegacyDrawing()
        {
            return chartsheet.legacyDrawing;
        }


        internal override void Write(Stream out1)
        {
            new ChartsheetDocument(this.chartsheet).Save(out1);
        }

        private static byte[] blankWorksheet()
        {
            MemoryStream out1 = new MemoryStream();
            try
            {
                new XSSFSheet().Write(out1);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            return out1.ToArray();
        }
    }


}
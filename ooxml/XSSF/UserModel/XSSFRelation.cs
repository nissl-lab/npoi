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
using NPOI.Util;
using System.Collections.Generic;
using System;
using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.Model;
using System.IO;
namespace NPOI.XSSF.UserModel
{

    /**
     *
     */
    public class XSSFRelation : POIXMLRelation
    {

        private static POILogger log = POILogFactory.GetLogger(typeof(XSSFRelation));

        /**
         * A map to lookup POIXMLRelation by its relation type
         */
        protected static Dictionary<String, XSSFRelation> _table = new Dictionary<String, XSSFRelation>();


        public static XSSFRelation WORKBOOK = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/workbook",
                "/xl/workbook.xml",
                null
        );
        public static XSSFRelation MACROS_WORKBOOK = new XSSFRelation(
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                PackageRelationshipTypes.CORE_DOCUMENT,
                "/xl/workbook.xml",
                null
        );
        public static XSSFRelation TEMPLATE_WORKBOOK = new XSSFRelation(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
                  PackageRelationshipTypes.CORE_DOCUMENT,
                  "/xl/workbook.xml",
                  null
        );
        public static XSSFRelation MACRO_TEMPLATE_WORKBOOK = new XSSFRelation(
                  "application/vnd.ms-excel.template.macroEnabled.main+xml",
                  PackageRelationshipTypes.CORE_DOCUMENT,
                  "/xl/workbook.xml",
                  null
        );
        public static XSSFRelation MACRO_ADDIN_WORKBOOK = new XSSFRelation(
                  "application/vnd.ms-excel.Addin.macroEnabled.main+xml",
                  PackageRelationshipTypes.CORE_DOCUMENT,
                  "/xl/workbook.xml",
                  null
        );

        public static XSSFRelation XLSB_BINARY_WORKBOOK = new XSSFRelation(
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
            PackageRelationshipTypes.CORE_DOCUMENT,
            "/xl/workbook.bin",
            null
        );
        public static XSSFRelation WORKSHEET = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "/xl/worksheets/sheet#.xml",
                typeof(XSSFSheet)
        );
        public static XSSFRelation CHARTSHEET = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
                "/xl/chartsheets/sheet#.xml",
                typeof(XSSFChartSheet)
        );
        public static XSSFRelation SHARED_STRINGS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                "/xl/sharedStrings.xml",
                typeof(SharedStringsTable)
        );
        public static XSSFRelation STYLES = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                PackageRelationshipTypes.STYLE_PART,
                "/xl/styles.xml",
                 typeof(StylesTable)
        );
        public static XSSFRelation DRAWINGS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.drawing+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                "/xl/drawings/drawing#.xml",
                typeof(XSSFDrawing)
        );
        public static XSSFRelation VML_DRAWINGS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.vmlDrawing",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                "/xl/drawings/vmlDrawing#.vml",
                typeof(XSSFVMLDrawing)
        );
        public static XSSFRelation CHART = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                "/xl/charts/chart#.xml",
                typeof(XSSFChart)
        );

        public static XSSFRelation CUSTOM_XML_MAPPINGS = new XSSFRelation(
                "application/xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps",
                "/xl/xmlMaps.xml",
                typeof(MapInfo)
        );

        public static XSSFRelation SINGLE_XML_CELLS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells",
                "/xl/tables/tableSingleCells#.xml",
                typeof(SingleXmlCells)
        );

        public static XSSFRelation TABLE = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
                "/xl/tables/table#.xml",
                typeof(XSSFTable)
        );

        public static XSSFRelation IMAGES = new XSSFRelation(
                null,
                PackageRelationshipTypes.IMAGE_PART,
                null,
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_EMF = new XSSFRelation(
                "image/x-emf",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.emf",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_WMF = new XSSFRelation(
                "image/x-wmf",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.wmf",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_PICT = new XSSFRelation(
                "image/pict",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.pict",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_JPEG = new XSSFRelation(
                "image/jpeg",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.jpeg",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_PNG = new XSSFRelation(
                "image/png",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.png",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_DIB = new XSSFRelation(
                "image/dib",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.dib",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_GIF = new XSSFRelation(
            "image/gif",
            PackageRelationshipTypes.IMAGE_PART,
            "/xl/media/image#.gif",
            typeof(XSSFPictureData)
        );

        public static XSSFRelation IMAGE_TIFF = new XSSFRelation(
                "image/tiff",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.tiff",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_EPS = new XSSFRelation(
                "image/x-eps",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.eps",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_BMP = new XSSFRelation(
                "image/x-ms-bmp",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.bmp",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation IMAGE_WPG = new XSSFRelation(
                "image/x-wpg",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.wpg",
                typeof(XSSFPictureData)
        );
        public static XSSFRelation SHEET_COMMENTS = new XSSFRelation(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                  "/xl/comments#.xml",
                  typeof(CommentsTable)
          );
        public static XSSFRelation SHEET_HYPERLINKS = new XSSFRelation(
                null,
                PackageRelationshipTypes.HYPERLINK_PART,
                null,
                null
        );
        public static XSSFRelation OLEEMBEDDINGS = new XSSFRelation(
                null,
                POIXMLDocument.OLE_OBJECT_REL_TYPE,
                null,
                null
        );
        public static XSSFRelation PACKEMBEDDINGS = new XSSFRelation(
                null,
                POIXMLDocument.PACK_OBJECT_REL_TYPE,
                null,
                null
        );

        public static XSSFRelation VBA_MACROS = new XSSFRelation(
                "application/vnd.ms-office.vbaProject",
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                "/xl/vbaProject.bin",
                null
        );
        public static XSSFRelation ACTIVEX_CONTROLS = new XSSFRelation(
                "application/vnd.ms-office.activeX+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
                "/xl/activeX/activeX#.xml",
                null
        );
        public static XSSFRelation ACTIVEX_BINS = new XSSFRelation(
                "application/vnd.ms-office.activeX",
                "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary",
                "/xl/activeX/activeX#.bin",
                null
        );
        public static XSSFRelation THEME = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.theme+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                "/xl/theme/theme#.xml",
                typeof(ThemesTable)
        );
        public static XSSFRelation CALC_CHAIN = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
                "/xl/calcChain.xml",
                typeof(CalculationChain)
        );

        public static XSSFRelation EXTERNAL_LINKS = new XSSFRelation(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
        "/xl/externalLinks/externalLink#.xmll",
        typeof(ExternalLinksTable)
        );

        public static XSSFRelation PRINTER_SETTINGS = new XSSFRelation(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
              "/xl/printerSettings/printerSettings#.bin",
              null
       );

        public static  XSSFRelation PIVOT_TABLE = new XSSFRelation(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
            "/xl/pivotTables/pivotTable#.xml",
            typeof(XSSFPivotTable)
        );
        public static  XSSFRelation PIVOT_CACHE_DEFINITION = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
                "/xl/pivotCache/pivotCacheDefinition#.xml",
                typeof(XSSFPivotCacheDefinition)
        );
        public static  XSSFRelation PIVOT_CACHE_RECORDS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
                "/xl/pivotCache/pivotCacheRecords#.xml",
                typeof(XSSFPivotCacheRecords)
        );

        private XSSFRelation(String type, String rel, String defaultName, Type cls) :
            base(type, rel, defaultName, cls)
        {
            if (cls != null && !_table.ContainsKey(rel))
                _table.Add(rel, this);
        }

        /**
         *  Fetches the InputStream to read the contents, based
         *  of the specified core part, for which we are defined
         *  as a suitable relationship
         */
        public Stream GetContents(PackagePart corePart)
        {
            PackageRelationshipCollection prc =
                corePart.GetRelationshipsByType(_relation);
            IEnumerator<PackageRelationship> it = prc.GetEnumerator();
            if (it.MoveNext())
            {
                PackageRelationship rel = it.Current;
                PackagePartName relName = PackagingUriHelper.CreatePartName(rel.TargetUri);
                PackagePart part = corePart.Package.GetPart(relName);
                return part.GetInputStream();
            }
            log.Log(POILogger.WARN, "No part " + _defaultName + " found");
            return null;
        }


        /**
         * Get POIXMLRelation by relation type
         *
         * @param rel relation type, for example,
         *    <code>http://schemas.openxmlformats.org/officeDocument/2006/relationships/image</code>
         * @return registered POIXMLRelation or null if not found
         */
        public static XSSFRelation GetInstance(String rel)
        {
            if (_table.ContainsKey(rel))
                return _table[rel];
            else
                return null;
        }

        /// <summary>
        /// Removes the relation from the internal table.
        /// Following readings of files will ignoring the removed relation.
        /// </summary>
        /// <param name="relation">Relation to remove</param>
        public static void RemoveRelation(XSSFRelation relation)
        {
            if (_table.ContainsKey(relation._relation))
            {
                _table.Remove(relation._relation);
            }
        }

        /// <summary>
        /// Adds the relation to the internal table.
        /// Following readings of files will process the given relation.
        /// </summary>
        /// <param name="relation">Relation to add</param>
        internal static void AddRelation(XSSFRelation relation)
        {
            if ((null != relation._type) && !_table.ContainsKey(relation._relation))
            {
                _table.Add(relation._relation, relation);
            }
        }
    }
}



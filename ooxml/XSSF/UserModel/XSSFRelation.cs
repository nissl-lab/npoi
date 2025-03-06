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
using NPOI.OpenXml4Net.OPC;
using NPOI.Util;
using NPOI.XSSF.Model;
using System;
using System.Collections.Generic;
using System.IO;
namespace NPOI.XSSF.UserModel
{

    /// <summary>
    /// Defines namespaces, content types and normal file names / naming
    /// patterns, for the well-known XSSF format parts. 
    /// </summary>
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
                , null
                , null
                , null
        );
        public static XSSFRelation MACROS_WORKBOOK = new XSSFRelation(
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                PackageRelationshipTypes.CORE_DOCUMENT,
                "/xl/workbook.xml",
                null
                , null
                , null
                , null
        );
        public static XSSFRelation TEMPLATE_WORKBOOK = new XSSFRelation(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
                  PackageRelationshipTypes.CORE_DOCUMENT,
                  "/xl/workbook.xml",
                  null
                , null
                , null
                , null
        );
        public static XSSFRelation MACRO_TEMPLATE_WORKBOOK = new XSSFRelation(
                  "application/vnd.ms-excel.template.macroEnabled.main+xml",
                  PackageRelationshipTypes.CORE_DOCUMENT,
                  "/xl/workbook.xml",
                  null
                , null
                , null
                , null
        );
        public static XSSFRelation MACRO_ADDIN_WORKBOOK = new XSSFRelation(
                  "application/vnd.ms-excel.Addin.macroEnabled.main+xml",
                  PackageRelationshipTypes.CORE_DOCUMENT,
                  "/xl/workbook.xml",
                  null
                , null
                , null
                , null
        );

        public static XSSFRelation XLSB_BINARY_WORKBOOK = new XSSFRelation(
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
            PackageRelationshipTypes.CORE_DOCUMENT,
            "/xl/workbook.bin",
            null
                , null
                , null
                , null
        );
        public static XSSFRelation WORKSHEET = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "/xl/worksheets/sheet#.xml",
                typeof(XSSFSheet)
            , null
            , (part) => new XSSFSheet(part)
            , () => new XSSFSheet()
        );
        public static XSSFRelation CHARTSHEET = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
                "/xl/chartsheets/sheet#.xml",
                typeof(XSSFChartSheet)
            , null
            , (part) => new XSSFChartSheet(part)
            , null
        );
        public static XSSFRelation SHARED_STRINGS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                "/xl/sharedStrings.xml",
                typeof(SharedStringsTable)
            , null
            , (part) => new SharedStringsTable(part)
            , () => new SharedStringsTable()
        );
        public static XSSFRelation STYLES = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                PackageRelationshipTypes.STYLE_PART,
                "/xl/styles.xml",
                 typeof(StylesTable)
            , null
            , (part) => new StylesTable(part)
            , () => new StylesTable()
        );
        public static XSSFRelation DRAWINGS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.drawing+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                "/xl/drawings/drawing#.xml",
                typeof(XSSFDrawing)
            , null
            , (part) => new XSSFDrawing(part)
            , () => new XSSFDrawing()
        );
        public static XSSFRelation VML_DRAWINGS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.vmlDrawing",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                "/xl/drawings/vmlDrawing#.vml",
                typeof(XSSFVMLDrawing)
            , null
            , (part) => new XSSFVMLDrawing(part)
            , () => new XSSFVMLDrawing()
        );
        public static XSSFRelation CHART = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                "/xl/charts/chart#.xml",
                typeof(XSSFChart)
            , null
            , (part) => new XSSFChart(part)
            , () => new XSSFChart()
        );

        public static XSSFRelation CUSTOM_XML_MAPPINGS = new XSSFRelation(
                "application/xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps",
                "/xl/xmlMaps.xml",
                typeof(MapInfo)
            , null
            , (part) => new MapInfo(part)
            , () => new MapInfo()
        );

        public static XSSFRelation SINGLE_XML_CELLS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells",
                "/xl/tables/tableSingleCells#.xml",
                typeof(SingleXmlCells)
            , null
            , (part) => new SingleXmlCells(part)
            , () => new SingleXmlCells()
        );

        public static XSSFRelation TABLE = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
                "/xl/tables/table#.xml",
                typeof(XSSFTable)
            , null
            , (part) => new XSSFTable(part)
            , () => new XSSFTable()
        );

        public static XSSFRelation IMAGES = new XSSFRelation(
                null,
                PackageRelationshipTypes.IMAGE_PART,
                null,
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_EMF = new XSSFRelation(
                "image/x-emf",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.emf",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_WMF = new XSSFRelation(
                "image/x-wmf",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.wmf",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_PICT = new XSSFRelation(
                "image/pict",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.pict",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_JPEG = new XSSFRelation(
                "image/jpeg",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.jpeg",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        /** JPG is an intentional duplicate of JPEG, to handle documents generated by other software. **/
        public static XSSFRelation IMAGE_JPG = new XSSFRelation(
                "image/jpg",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.jpg",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_PNG = new XSSFRelation(
                "image/png",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.png",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_DIB = new XSSFRelation(
                "image/dib",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.dib",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_GIF = new XSSFRelation(
            "image/gif",
            PackageRelationshipTypes.IMAGE_PART,
            "/xl/media/image#.gif",
            typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );

        public static XSSFRelation IMAGE_TIFF = new XSSFRelation(
                "image/tiff",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.tiff",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_EPS = new XSSFRelation(
                "image/x-eps",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.eps",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_BMP = new XSSFRelation(
                "image/x-ms-bmp",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.bmp",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation IMAGE_WPG = new XSSFRelation(
                "image/x-wpg",
                PackageRelationshipTypes.IMAGE_PART,
                "/xl/media/image#.wpg",
                typeof(XSSFPictureData)
            , null
            , (part) => new XSSFPictureData(part)
            , () => new XSSFPictureData()
        );
        public static XSSFRelation SHEET_COMMENTS = new XSSFRelation(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                  "/xl/comments#.xml",
                  typeof(CommentsTable)
            , null
            , (part) => new CommentsTable(part)
            , () => new CommentsTable()
          );
        public static XSSFRelation SHEET_HYPERLINKS = new XSSFRelation(
                null,
                PackageRelationshipTypes.HYPERLINK_PART,
                null,
                null
            , null
            , null
            , null
        );
        public static XSSFRelation OLEEMBEDDINGS = new XSSFRelation(
                null,
                POIXMLDocument.OLE_OBJECT_REL_TYPE,
                null,
                null
            , null
            , null
            , null
        );
        public static XSSFRelation PACKEMBEDDINGS = new XSSFRelation(
                null,
                POIXMLDocument.PACK_OBJECT_REL_TYPE,
                null,
                null
            , null
            , null
            , null
        );

        public static XSSFRelation VBA_MACROS = new XSSFRelation(
                "application/vnd.ms-office.vbaProject",
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                "/xl/vbaProject.bin",
                typeof(XSSFVBAPart)
            , null
            , (part) => new XSSFVBAPart(part)
            , () => new XSSFVBAPart()
        );
        public static XSSFRelation ACTIVEX_CONTROLS = new XSSFRelation(
                "application/vnd.ms-office.activeX+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
                "/xl/activeX/activeX#.xml",
                null
            , null
            , null
            , null
        );
        public static XSSFRelation ACTIVEX_BINS = new XSSFRelation(
                "application/vnd.ms-office.activeX",
                "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary",
                "/xl/activeX/activeX#.bin",
                null
            , null
            , null
            , null
        );
        public static XSSFRelation THEME = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.theme+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                "/xl/theme/theme#.xml",
                typeof(ThemesTable)
            , null
            , (part) => new ThemesTable(part)
            , () => new ThemesTable()
        );
        public static XSSFRelation CALC_CHAIN = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
                "/xl/calcChain.xml",
                typeof(CalculationChain)
            , null
            , (part) => new CalculationChain(part)
            , () => new CalculationChain()
        );

        public static XSSFRelation EXTERNAL_LINKS = new XSSFRelation(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
        "/xl/externalLinks/externalLink#.xmll",
        typeof(ExternalLinksTable)
            , null
            , (part) => new ExternalLinksTable(part)
            , () => new ExternalLinksTable()
        );

        public static XSSFRelation PRINTER_SETTINGS = new XSSFRelation(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
              "/xl/printerSettings/printerSettings#.bin",
              null
            , null
            , null
            , null
       );

        public static  XSSFRelation PIVOT_TABLE = new XSSFRelation(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
            "/xl/pivotTables/pivotTable#.xml",
            typeof(XSSFPivotTable)
            , null
            , (part) => new XSSFPivotTable(part)
            , () => new XSSFPivotTable()
        );
        public static  XSSFRelation PIVOT_CACHE_DEFINITION = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
                "/xl/pivotCache/pivotCacheDefinition#.xml",
                typeof(XSSFPivotCacheDefinition)
            , null
            , (part) => new XSSFPivotCacheDefinition(part)
            , () => new XSSFPivotCacheDefinition()
        );
        public static  XSSFRelation PIVOT_CACHE_RECORDS = new XSSFRelation(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
                "/xl/pivotCache/pivotCacheRecords#.xml",
                typeof(XSSFPivotCacheRecords)
            , null
            , (part) => new XSSFPivotCacheRecords(part)
            , () => new XSSFPivotCacheRecords()
        );

        public static XSSFRelation CTRL_PROP_RECORDS = new XSSFRelation(
            null,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp",
            "/xl/ctrlProps/ctrlProp#.xml",
            null
            , null
            , null
            , null
        );

        public static XSSFRelation CUSTOM_PROPERTIES = new XSSFRelation(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty",
            "/xl/customProperty#.bin",
            null
            , null
            , null
            , null
    );

        public static String NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public static String NS_DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static String NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        private XSSFRelation(String type, String rel, String defaultName, Type cls
            , Func<POIXMLDocumentPart, PackagePart, POIXMLDocumentPart> createPartWithParent
            , Func<PackagePart, POIXMLDocumentPart> createPart
            , Func<POIXMLDocumentPart> createInstance) :
            base(type, rel, defaultName, cls, createPartWithParent, createPart, createInstance)
        {
            _table[rel] = this; 
        }

        /**
         *  Fetches the InputStream to read the contents, based
         *  of the specified core part, for which we are defined
         *  as a suitable relationship
         */
        public Stream GetContents(PackagePart corePart)
        {
            PackageRelationshipCollection prc =
                corePart.GetRelationshipsByType(Relation);
            IEnumerator<PackageRelationship> it = prc.GetEnumerator();
            if (it.MoveNext())
            {
                PackageRelationship rel = it.Current;
                PackagePartName relName = PackagingUriHelper.CreatePartName(rel.TargetUri);
                PackagePart part = corePart.Package.GetPart(relName);
                return part.GetInputStream();
            }
            log.Log(POILogger.WARN, "No part " + DefaultFileName + " found");
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
            if (_table.TryGetValue(rel, out XSSFRelation result))
            {
                return result;
            }

            return null;
        }

        /// <summary>
        /// Removes the relation from the internal table.
        /// Following readings of files will ignoring the removed relation.
        /// </summary>
        /// <param name="relation">Relation to remove</param>
        public static bool RemoveRelation(XSSFRelation relation)
        {
            return _table.Remove(relation.Relation);
        }

        /// <summary>
        /// Adds the relation to the internal table.
        /// Following readings of files will process the given relation.
        /// </summary>
        /// <param name="relation">Relation to add</param>
        internal static void AddRelation(XSSFRelation relation)
        {
            if (relation.ContentType is not null && !_table.ContainsKey(relation.Relation))
            {
                _table.Add(relation.Relation, relation);
            }
        }
    }
}



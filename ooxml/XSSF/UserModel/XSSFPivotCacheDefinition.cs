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
    using System.IO;
    using System.Xml;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using System;
    using NPOI.SS;

    public class XSSFPivotCacheDefinition : POIXMLDocumentPart
    {

        private CT_PivotCacheDefinition ctPivotCacheDefinition;
        
        public XSSFPivotCacheDefinition()
            : base()
        {

            ctPivotCacheDefinition = new CT_PivotCacheDefinition();
            CreateDefaultValues();
        }

        /**
        * Creates an XSSFPivotCacheDefintion representing the given package part and relationship.
        * Should only be called when Reading in an existing file.
        *
        * @param part - The package part that holds xml data representing this pivot cache defInition.
        * @param rel - the relationship of the given package part in the underlying OPC package
        */

        internal XSSFPivotCacheDefinition(PackagePart part)
            : base(part)
        {

            ReadFrom(part.GetInputStream());
        }

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        protected XSSFPivotCacheDefinition(PackagePart part, PackageRelationship rel)
            : this(part)
        {

        }
        public void ReadFrom(Stream is1)
        {
            try
            {
                //XmlOptions options  = new XmlOptions(DEFAULT_XML_OPTIONS);
                ////Removing root element
                //options.LoadReplaceDocumentElement=(/*setter*/null);
                XmlDocument xmldoc = ConvertStreamToXml(is1);
                ctPivotCacheDefinition = CT_PivotCacheDefinition.Parse(xmldoc.DocumentElement, NamespaceManager);
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }



        public CT_PivotCacheDefinition GetCTPivotCacheDefinition()
        {
            return ctPivotCacheDefinition;
        }


        private void CreateDefaultValues()
        {
            ctPivotCacheDefinition.createdVersion = (byte)XSSFPivotTable.CREATED_VERSION;
            ctPivotCacheDefinition.minRefreshableVersion = (byte)XSSFPivotTable.MIN_REFRESHABLE_VERSION;
            ctPivotCacheDefinition.refreshedVersion = (byte)XSSFPivotTable.UPDATED_VERSION;
            ctPivotCacheDefinition.refreshedBy = (/*setter*/"NPOI");
            ctPivotCacheDefinition.refreshedDate = DateTime.Now.ToOADate();
            ctPivotCacheDefinition.refreshOnLoad = (/*setter*/true);
        }



        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //Sets the pivotCacheDefInition tag
            //xmlOptions.SetSaveSyntheticDocumentElement(new QName(CTPivotCacheDefInition.type.Name.
            //         GetNamespaceURI(), "pivotCacheDefInition"));

            //// ensure the fields have names
            //if (ctPivotCacheDefinition.cacheFields != null)
            //{
            //    CT_CacheFields cFields = ctPivotCacheDefinition.cacheFields;
            //    foreach (CT_CacheField cf in cFields.cacheField)
            //    {
            //        if (cf.name == null || cf.name.Equals(""))
            //        {
            //            cf.name = "A";
            //        }
            //    }
            //}

            ctPivotCacheDefinition.Save(out1);
            out1.Close();
        }

        /**
         * Find the 2D base data area for the pivot table, either from its direct reference or named table/range.
         * @return AreaReference representing the current area defined by the pivot table
         * @ if the ref1 attribute is not contiguous or the name attribute is not found.
         */

        public AreaReference GetPivotArea(IWorkbook wb)
        {
            CT_WorksheetSource wsSource = ctPivotCacheDefinition.cacheSource.worksheetSource;

            String ref1 = wsSource.@ref;
            String name = wsSource.name;

            if (ref1 == null && name == null)
                throw new ArgumentException("Pivot cache must reference an area, named range, or table.");

            // this is the XML format, so tell the reference that.
            if (ref1 != null) return new AreaReference(ref1, SpreadsheetVersion.EXCEL2007);

            if (name != null)
            {
                // named range or table?
                IName range = wb.GetName(name);
                if (range != null) return new AreaReference(range.RefersToFormula, SpreadsheetVersion.EXCEL2007);
                // not a named range, check for a table.
                // do this second, as tables are sheet-specific, but named ranges are not, and may not have a sheet name given.
                XSSFSheet sheet = (XSSFSheet)wb.GetSheet(wsSource.sheet);
                foreach (XSSFTable table in sheet.GetTables())
                {
                    if (table.Name.Equals(name))
                    { //case-sensitive?
                        return new AreaReference(table.StartCellReference, table.EndCellReference);
                    }
                }
            }

            throw new ArgumentException("Name '" + name + "' was not found.");
        }


        /**
         * Generates a cache field for each column in the reference area for the pivot table.
         * @param sheet The sheet where the data i collected from
         */

        protected internal void CreateCacheFields(ISheet sheet)
        {
            //Get values for start row, start and end column
            AreaReference ar = GetPivotArea(sheet.Workbook);
            CellReference firstCell = ar.FirstCell;
            CellReference lastCell = ar.LastCell;
            int columnStart = firstCell.Col;
            int columnEnd = lastCell.Col;
            IRow row = sheet.GetRow(firstCell.Row);
            CT_CacheFields cFields;
            if (ctPivotCacheDefinition.cacheFields != null)
            {
                cFields = ctPivotCacheDefinition.cacheFields;
            }
            else
            {
                cFields = ctPivotCacheDefinition.AddNewCacheFields();
            }
            //For each column, create a cache field and give it en empty sharedItems
            for (int i = columnStart; i <= columnEnd; i++)
            {
                CT_CacheField cf = cFields.AddNewCacheField();
                if (i == columnEnd)
                {
                    cFields.count = (/*setter*/cFields.SizeOfCacheFieldArray());
                }
                //General number format
                cf.numFmtId = (/*setter*/0);
                ICell cell = row.GetCell(i);
                cell.SetCellType(CellType.String);
                String stringCellValue = cell.StringCellValue;
                cf.name = stringCellValue;
                cf.AddNewSharedItems();
            }
        }
    }
}
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
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.OpenXmlFormats.Spreadsheet;


    public class XSSFPivotTable : POIXMLDocumentPart
    {

        protected internal static short CREATED_VERSION = 3;
        protected internal static short MIN_REFRESHABLE_VERSION = 3;
        protected internal static short UPDATED_VERSION = 3;

        private CT_PivotTableDefinition pivotTableDefinition;
        private XSSFPivotCacheDefinition pivotCacheDefinition;
        private XSSFPivotCache pivotCache;
        private XSSFPivotCacheRecords pivotCacheRecords;
        private ISheet parentSheet;
        private ISheet dataSheet;


        public XSSFPivotTable()
            : base()
        {
            pivotTableDefinition = new CT_PivotTableDefinition();
            pivotCache = new XSSFPivotCache();
            pivotCacheDefinition = new XSSFPivotCacheDefinition();
            pivotCacheRecords = new XSSFPivotCacheRecords();
        }

        /**
        * Creates an XSSFPivotTable representing the given package part and relationship.
        * Should only be called when Reading in an existing file.
        *
        * @param part - The package part that holds xml data representing this pivot table.
        * @param rel - the relationship of the given package part in the underlying OPC package
        */

        protected XSSFPivotTable(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            ReadFrom(part.GetInputStream());
        }


        public void ReadFrom(Stream is1)
        {
            try
            {
                //XmlOptions options  = new XmlOptions(DEFAULT_XML_OPTIONS);
                //Removing root element
                //options.LoadReplaceDocumentElement=(/*setter*/null);
                XmlDocument xmlDoc = ConvertStreamToXml(is1);
                pivotTableDefinition = CT_PivotTableDefinition.Parse(xmlDoc.DocumentElement, NamespaceManager);
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }


        public void SetPivotCache(XSSFPivotCache pivotCache)
        {
            this.pivotCache = pivotCache;
        }


        public XSSFPivotCache GetPivotCache()
        {
            return pivotCache;
        }


        public ISheet GetParentSheet()
        {
            return parentSheet;
        }


        public void SetParentSheet(XSSFSheet parentSheet)
        {
            this.parentSheet = parentSheet;
        }



        public CT_PivotTableDefinition GetCTPivotTableDefinition()
        {
            return pivotTableDefinition;
        }



        public void SetCTPivotTableDefinition(CT_PivotTableDefinition pivotTableDefinition)
        {
            this.pivotTableDefinition = pivotTableDefinition;
        }


        public XSSFPivotCacheDefinition GetPivotCacheDefinition()
        {
            return pivotCacheDefinition;
        }


        public void SetPivotCacheDefinition(XSSFPivotCacheDefinition pivotCacheDefinition)
        {
            this.pivotCacheDefinition = pivotCacheDefinition;
        }


        public XSSFPivotCacheRecords GetPivotCacheRecords()
        {
            return pivotCacheRecords;
        }


        public void SetPivotCacheRecords(XSSFPivotCacheRecords pivotCacheRecords)
        {
            this.pivotCacheRecords = pivotCacheRecords;
        }


        public ISheet GetDataSheet()
        {
            return dataSheet;
        }


        private void SetDataSheet(ISheet dataSheet)
        {
            this.dataSheet = dataSheet;
        }



        protected internal override void Commit()
        {
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //Sets the pivotTableDefInition tag
            //xmlOptions.SetSaveSyntheticDocumentElement(new QName(CTPivotTableDefinition.type.Name.
            //        GetNamespaceURI(), "pivotTableDefinition"));
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            pivotTableDefinition.Save(out1);
            out1.Close();
        }

        /**
         * Set default values for the table defInition.
         */

        protected internal void SetDefaultPivotTableDefinition()
        {
            //Not more than one until more Created
            pivotTableDefinition.multipleFieldFilters = (/*setter*/false);
            //Indentation increment for compact rows
            pivotTableDefinition.indent = (/*setter*/0);
            //The pivot version which Created the pivot cache Set to default value
            pivotTableDefinition.createdVersion = (byte)CREATED_VERSION;
            //Minimun version required to update the pivot cache
            pivotTableDefinition.minRefreshableVersion = (byte)MIN_REFRESHABLE_VERSION;
            //Version of the application which "updated the spreadsheet last"
            pivotTableDefinition.updatedVersion = (byte)UPDATED_VERSION;
            //Titles Shown at the top of each page when printed
            pivotTableDefinition.itemPrintTitles = (/*setter*/true);
            //Set autoformat properties
            pivotTableDefinition.useAutoFormatting = (/*setter*/true);
            pivotTableDefinition.applyNumberFormats = (/*setter*/false);
            pivotTableDefinition.applyWidthHeightFormats = (/*setter*/true);
            pivotTableDefinition.applyAlignmentFormats = (/*setter*/false);
            pivotTableDefinition.applyPatternFormats = (/*setter*/false);
            pivotTableDefinition.applyFontFormats = (/*setter*/false);
            pivotTableDefinition.applyBorderFormats = (/*setter*/false);
            pivotTableDefinition.cacheId = (/*setter*/pivotCache.GetCTPivotCache().cacheId);
            pivotTableDefinition.name = (/*setter*/"PivotTable" + pivotTableDefinition.cacheId);
            pivotTableDefinition.dataCaption = (/*setter*/"Values");

            //Set the default style for the pivot table
            CT_PivotTableStyle style = pivotTableDefinition.AddNewPivotTableStyleInfo();
            style.name = (/*setter*/"PivotStyleLight16");
            style.showLastColumn = (/*setter*/true);
            style.showColStripes = (/*setter*/false);
            style.showRowStripes = (/*setter*/false);
            style.showColHeaders = (/*setter*/true);
            style.showRowHeaders = (/*setter*/true);
        }

        protected AreaReference GetPivotArea()
        {
            AreaReference pivotArea = new AreaReference(GetPivotCacheDefinition().
                    GetCTPivotCacheDefInition().cacheSource.worksheetSource.@ref);
            return pivotArea;
        }

        /**
         * Add a row label using data from the given column.
         * @param columnIndex the index of the column to be used as row label.
         */

        public void AddRowLabel(int columnIndex)
        {
            AreaReference pivotArea = GetPivotArea();
            int lastRowIndex = pivotArea.LastCell.Row - pivotArea.FirstCell.Row;
            int lastColIndex = pivotArea.LastCell.Col - pivotArea.FirstCell.Col;

            if (columnIndex > lastColIndex)
            {
                throw new IndexOutOfRangeException();
            }
            CT_PivotFields pivotFields = pivotTableDefinition.pivotFields;

            CT_PivotField pivotField = new CT_PivotField();
            CT_Items items = pivotField.AddNewItems();

            pivotField.axis = (/*setter*/ST_Axis.axisRow);
            pivotField.showAll = (/*setter*/false);
            for (int i = 0; i <= lastRowIndex; i++)
            {
                items.AddNewItem().t = (/*setter*/ST_ItemType.@default);
            }
            items.count = (/*setter*/items.SizeOfItemArray());
            pivotFields.SetPivotFieldArray(columnIndex, pivotField);

            CT_RowFields rowFields;
            if (pivotTableDefinition.rowFields != null)
            {
                rowFields = pivotTableDefinition.rowFields;
            }
            else
            {
                rowFields = pivotTableDefinition.AddNewRowFields();
            }

            rowFields.AddNewField().x = (/*setter*/columnIndex);
            rowFields.count = (/*setter*/rowFields.SizeOfFieldArray());
        }

        public IList<int> GetRowLabelColumns()
        {
            if (pivotTableDefinition.rowFields != null)
            {
                List<int> columnIndexes = new List<int>();
                foreach (CT_Field f in pivotTableDefinition.rowFields.GetFieldArray())
                {
                    columnIndexes.Add(f.x);
                }
                return columnIndexes;
            }
            else
            {
                return new List<int>();
            }
        }

        /**
         * Add a column label using data from the given column and specified function
         * @param columnIndex the index of the column to be used as column label.
         * @param function the function to be used on the data
         * The following functions exists:
         * Sum, Count, Average, Max, Min, Product, Count numbers, StdDev, StdDevp, Var, Varp
         * @param valueFieldName the name of pivot table value field
         */

        public void AddColumnLabel(DataConsolidateFunction function, int columnIndex, String valueFieldName)
        {
            AreaReference pivotArea = GetPivotArea();
            int lastColIndex = pivotArea.LastCell.Col - pivotArea.FirstCell.Col;

            if (columnIndex > lastColIndex && columnIndex < 0)
            {
                throw new IndexOutOfRangeException();
            }

            AddDataColumn(columnIndex, true);
            AddDataField(function, columnIndex, valueFieldName);

            // colfield should be Added for the second one.
            if (pivotTableDefinition.dataFields.count == 2)
            {
                CT_ColFields colFields;
                if (pivotTableDefinition.colFields != null)
                {
                    colFields = pivotTableDefinition.colFields;
                }
                else
                {
                    colFields = pivotTableDefinition.AddNewColFields();
                }
                colFields.AddNewField().x = (/*setter*/-2);
                colFields.count = (/*setter*/colFields.SizeOfFieldArray());
            }
        }

        /**
         * Add a column label using data from the given column and specified function
         * @param columnIndex the index of the column to be used as column label.
         * @param function the function to be used on the data
         * The following functions exists:
         * Sum, Count, Average, Max, Min, Product, Count numbers, StdDev, StdDevp, Var, Varp
         */

        public void AddColumnLabel(DataConsolidateFunction function, int columnIndex)
        {
            AddColumnLabel(function, columnIndex, function.Name);
        }

        /**
         * Add data field with data from the given column and specified function.
         * @param function the function to be used on the data
         * @param index the index of the column to be used as column label.
         * The following functions exists:
         * Sum, Count, Average, Max, Min, Product, Count numbers, StdDev, StdDevp, Var, Varp
         * @param valueFieldName the name of pivot table value field
         */

        private void AddDataField(DataConsolidateFunction function, int columnIndex, String valueFieldName)
        {
            AreaReference pivotArea = GetPivotArea();
            int lastColIndex = pivotArea.LastCell.Col - pivotArea.FirstCell.Col;

            if (columnIndex > lastColIndex && columnIndex < 0)
            {
                throw new IndexOutOfRangeException();
            }
            CT_DataFields dataFields;
            if (pivotTableDefinition.dataFields != null)
            {
                dataFields = pivotTableDefinition.dataFields;
            }
            else
            {
                dataFields = pivotTableDefinition.AddNewDataFields();
            }
            CT_DataField dataField = dataFields.AddNewDataField();
            dataField.subtotal = (ST_DataConsolidateFunction)(function.Value);
            ICell cell = GetDataSheet().GetRow(pivotArea.FirstCell.Row).GetCell(columnIndex);
            cell.SetCellType(CellType.String);
            dataField.name = (/*setter*/valueFieldName);
            dataField.fld = (uint)columnIndex;
            dataFields.count = (/*setter*/dataFields.SizeOfDataFieldArray());
        }

        /**
         * Add column Containing data from the referenced area.
         * @param columnIndex the index of the column Containing the data
         * @param isDataField true if the data should be displayed in the pivot table.
         */

        public void AddDataColumn(int columnIndex, bool isDataField)
        {
            AreaReference pivotArea = GetPivotArea();
            int lastColIndex = pivotArea.LastCell.Col - pivotArea.FirstCell.Col;
            if (columnIndex > lastColIndex && columnIndex < 0)
            {
                throw new IndexOutOfRangeException();
            }
            CT_PivotFields pivotFields = pivotTableDefinition.pivotFields;
            CT_PivotField pivotField = new CT_PivotField();

            pivotField.dataField = (/*setter*/isDataField);
            pivotField.showAll = (/*setter*/false);
            pivotFields.SetPivotFieldArray(columnIndex, pivotField);
        }

        /**
         * Add filter for the column with the corresponding index and cell value
         * @param columnIndex index of column to filter on
         */

        public void AddReportFilter(int columnIndex)
        {
            AreaReference pivotArea = GetPivotArea();
            int lastColIndex = pivotArea.LastCell.Col - pivotArea.FirstCell.Col;
            int lastRowIndex = pivotArea.LastCell.Row - pivotArea.FirstCell.Row;

            if (columnIndex > lastColIndex && columnIndex < 0)
            {
                throw new IndexOutOfRangeException();
            }
            CT_PivotFields pivotFields = pivotTableDefinition.pivotFields;

            CT_PivotField pivotField = new CT_PivotField();
            CT_Items items = pivotField.AddNewItems();

            pivotField.axis = (/*setter*/ST_Axis.axisPage);
            pivotField.showAll = (/*setter*/false);
            for (int i = 0; i <= lastRowIndex; i++)
            {
                items.AddNewItem().t = (/*setter*/ST_ItemType.@default);
            }
            items.count = (/*setter*/items.SizeOfItemArray());
            pivotFields.SetPivotFieldArray(columnIndex, pivotField);

            CT_PageFields pageFields;
            if (pivotTableDefinition.pageFields != null)
            {
                pageFields = pivotTableDefinition.pageFields;
                //Another filter has already been Created
                pivotTableDefinition.multipleFieldFilters = (/*setter*/true);
            }
            else
            {
                pageFields = pivotTableDefinition.AddNewPageFields();
            }
            CT_PageField pageField = pageFields.AddNewPageField();
            pageField.hier = (/*setter*/-1);
            pageField.fld = (/*setter*/columnIndex);

            pageFields.count = (/*setter*/pageFields.SizeOfPageFieldArray());
            pivotTableDefinition.location.colPageCount = (/*setter*/pageFields.count);
        }

        /**
         * Creates cacheSource and workSheetSource for pivot table and Sets the source reference as well assets the location of the pivot table
         * @param source Source for data for pivot table
         * @param position Position for pivot table in sheet
         * @param sourceSheet Sheet where the source will be collected from
         */

        protected internal void CreateSourceReferences(AreaReference source, CellReference position, ISheet sourceSheet)
        {
            //Get cell one to the right and one down from position, add both to AreaReference and Set pivot table location.
            AreaReference destination = new AreaReference(position, new CellReference(position.Row + 1, position.Col + 1));

            CT_Location location;
            if (pivotTableDefinition.location == null)
            {
                location = pivotTableDefinition.AddNewLocation();
                location.firstDataCol = (/*setter*/1);
                location.firstDataRow = (/*setter*/1);
                location.firstHeaderRow = (/*setter*/1);
            }
            else
            {
                location = pivotTableDefinition.location;
            }
            location.@ref = (/*setter*/destination.FormatAsString());
            pivotTableDefinition.location = (/*setter*/location);

            //Set source for the pivot table
            CT_PivotCacheDefinition cacheDef = GetPivotCacheDefinition().GetCTPivotCacheDefInition();
            CT_CacheSource cacheSource = cacheDef.AddNewCacheSource();
            cacheSource.type = (/*setter*/ST_SourceType.worksheet);
            CT_WorksheetSource worksheetSource = cacheSource.AddNewWorksheetSource();
            worksheetSource.sheet = (/*setter*/sourceSheet.SheetName);
            SetDataSheet(sourceSheet);

            String[] firstCell = source.FirstCell.CellRefParts;
            String[] lastCell = source.LastCell.CellRefParts;
            worksheetSource.@ref = (/*setter*/firstCell[2] + firstCell[1] + ':' + lastCell[2] + lastCell[1]);
        }


        protected internal void CreateDefaultDataColumns()
        {
            CT_PivotFields pivotFields;
            if (pivotTableDefinition.pivotFields != null)
            {
                pivotFields = pivotTableDefinition.pivotFields;
            }
            else
            {
                pivotFields = pivotTableDefinition.AddNewPivotFields();
            }
            AreaReference sourceArea = GetPivotArea();
            int firstColumn = sourceArea.FirstCell.Col;
            int lastColumn = sourceArea.LastCell.Col;
            CT_PivotField pivotField;
            for (int i = 0; i <= lastColumn - firstColumn; i++)
            {
                pivotField = pivotFields.AddNewPivotField();
                pivotField.dataField = (/*setter*/false);
                pivotField.showAll = (/*setter*/false);
            }
            pivotFields.count = (/*setter*/pivotFields.SizeOfPivotFieldArray());
        }
    }

}
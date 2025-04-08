/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */


using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.Openxml4Net.Exceptions;
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.OpenXmlFormats.Dml.Chart;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.Util.Optional;
    using NPOI.XDDF.UserModel;
    using NPOI.XDDF.UserModel.Text;
    using NPOI.XSSF.UserModel;
    using System.Xml;

    public abstract class XDDFChart : POIXMLDocumentPart, ITextContainer
    {
        /// <summary>
        /// Underlying workbook
        /// </summary>
        private XSSFWorkbook workbook;

        private int chartIndex = 0;

        private POIXMLDocumentPart documentPart = null;

        protected List<XDDFChartAxis> axes = new List<XDDFChartAxis>();

        /// <summary>
        /// Root element of the Chart part
        /// </summary>
        protected  CT_ChartSpace chartSpace;

        /// <summary>
        /// Chart element in the chart space
        /// </summary>
        protected  CT_Chart chart;

        /// <summary>
        /// Construct a chart.
        /// </summary>
        protected XDDFChart() : base()
        {
            chartSpace = new CT_ChartSpace();
            chart = chartSpace.AddNewChart();
            chart.AddNewPlotArea();
        }

        /// <summary>
        /// Construct a DrawingML chart from a package part.
        /// </summary>
        /// <param name="part">
        /// the package part holding the chart data, the content type must
        /// be
        /// <c>application/vnd.Openxmlformats-officedocument.Drawingml.Chart+xml</c>
        /// </param>
        /// <remarks>
        /// @since POI 3.14-Beta1
        /// </remarks>
        protected XDDFChart(PackagePart part)
                : base(part)
        {
            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            chartSpace = ChartSpaceDocument.Parse(xml,
                POIXMLDocumentPart.NamespaceManager).GetChartSpace();
            chart = chartSpace.chart;
        }

        /// <summary>
        /// Return the underlying CTChartSpace bean, the root element of the Chart
        /// part.
        /// </summary>
        /// <returns>the underlying CTChartSpace bean</returns>
        public CT_ChartSpace GetCTChartSpace()
        {
            return chartSpace;
        }

        /// <summary>
        /// Return the underlying CTChart bean, within the Chart Space
        /// </summary>
        /// <returns>the underlying CTChart bean</returns>
        public CT_Chart GetCTChart()
        {
            return chart;

        }

        /// <summary>
        /// Return the underlying CTPlotArea bean, within the Chart
        /// </summary>
        /// <returns>the underlying CTPlotArea bean</returns>
        protected CT_PlotArea GetCTPlotArea()
        {
            return chart.plotArea;
        }

        /// <summary>
        /// </summary>
        /// <returns>true if only visible cells will be present on the chart, false
        /// otherwise
        /// </returns>
        public bool IsPlotOnlyVisibleCells()
        {
            if(chart.IsSetPlotVisOnly())
            {
                return chart.plotVisOnly.val == 1;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="only">
        /// a flag specifying if only visible cells should be present on
        /// the chart
        /// </param>
        public void SetPlotOnlyVisibleCells(bool only)
        {
            if(!chart.IsSetPlotVisOnly())
            {
                chart.plotVisOnly = new OpenXmlFormats.Dml.Chart.CT_Boolean();
            }
            chart.plotVisOnly.val = only ? 1 : 0;
        }

        public void SetFloor(int thickness)
        {
            if(!chart.IsSetFloor())
            {
                chart.floor = new CT_Surface();
            }
            chart.floor.thickness.val = (uint) thickness;
        }

        public void SetBackWall(int thickness)
        {
            if(!chart.IsSetBackWall())
            {
                chart.backWall = new CT_Surface();
            }
            chart.backWall.thickness.val = (uint) thickness;
        }

        public void SetSideWall(int thickness)
        {
            if(!chart.IsSetSideWall())
            {
                chart.sideWall = new CT_Surface();
            }
            chart.sideWall.thickness.val = (uint) thickness;
        }

        public void SetAutoTitleDeleted(bool deleted)
        {
            if(!chart.IsSetAutoTitleDeleted())
            {
                chart.autoTitleDeleted = new OpenXmlFormats.Dml.Chart.CT_Boolean();
            }
            chart.autoTitleDeleted.val = deleted ? 1 : 0;
        }

        /// <summary>
        /// </summary>
        ///
        /// <remarks>
        /// @since 4.0.1
        /// </remarks>
        public void DisplayBlanksAs(DisplayBlanks? as1)
        {
            if(!as1.HasValue)
            {
                if(chart.IsSetDispBlanksAs())
                {
                    chart.UnsetDispBlanksAs();
                }
            }
            else
            {
                if(chart.IsSetDispBlanksAs())
                {
                    chart.dispBlanksAs.val = as1.Value.ToST_DispBlanksAs();
                }
                else
                {
                    chart.AddNewDispBlanksAs().val = as1.Value.ToST_DispBlanksAs();
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <remarks>
        /// @since 4.0.1
        /// </remarks>
        public Boolean? TitleOverlay
        {
            get
            {
                if(chart.IsSetTitle())
                {
                    CT_Title title = chart.title;
                    if(title.IsSetOverlay())
                    {
                        return title.overlay.val == 1;
                    }
                }
                return null;
            }
            set
            {
                if(!chart.IsSetTitle())
                {
                    chart.AddNewTitle();
                }
                new XDDFTitle(this, chart.title).SetOverlay(value);
            }
        }

        /// <summary>
        /// Sets the title text as a static string.
        /// </summary>
        /// <param name="text">
        /// to use as new title
        /// </param>
        /// <remarks>
        /// @since 4.0.1
        /// </remarks>
        public void SetTitleText(string text)
        {
            if(!chart.IsSetTitle())
            {
                chart.AddNewTitle();
            }
            new XDDFTitle(this, chart.title).SetText(text);
        }

        /// <summary>
        /// </summary>
        /// <remarks>
        /// @since 4.0.1
        /// </remarks>
        public XDDFTitle GetTitle()
        {
            if(chart.IsSetTitle())
            {
                return new XDDFTitle(this, chart.title);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the chart title body if there is one, i.e. title is Set and is not a
        /// formula.
        /// </summary>
        /// <returns>text body or null, if title is a formula or no title is Set.</returns>
        public XDDFTextBody GetFormattedTitle()
        {
            if(!chart.IsSetTitle())
            {
                return null;
            }
            return new XDDFTitle(this, chart.title).Body;
        }
        public Option<R> FindDefinedParagraphProperty<R>(
                    Func<CT_TextParagraphProperties, Boolean> isSet,
                    Func<CT_TextParagraphProperties, R> Getter) where R : class
        {
            return Option<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public Option<R> FindDefinedRunProperty<R>(
                Func<CT_TextCharacterProperties, Boolean> isSet,
                Func<CT_TextCharacterProperties, R> Getter) where R : class
        {
            return Option<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public ValueOption<R> FindDefinedParagraphValueProperty<R>(
                Func<CT_TextParagraphProperties, Boolean> isSet,
                Func<CT_TextParagraphProperties, R> Getter) where R : struct
        {
            return ValueOption<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public ValueOption<R> FindDefinedRunValueProperty<R>(
                Func<CT_TextCharacterProperties, Boolean> isSet,
                Func<CT_TextCharacterProperties, R> Getter) where R : struct
        {
            return ValueOption<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public XDDFShapeProperties GetOrAddShapeProperties()
        {
            CT_PlotArea plotArea = GetCTPlotArea();
            CT_ShapeProperties properties;
            if(plotArea.IsSetSpPr())
            {
                properties = plotArea.spPr;
            }
            else
            {
                properties = plotArea.AddNewSpPr();
            }
            return new XDDFShapeProperties(properties);
        }

        public void DeleteShapeProperties()
        {
            if(GetCTPlotArea().IsSetSpPr())
            {
                GetCTPlotArea().UnsetSpPr();
            }
        }

        public XDDFChartLegend GetOrAddLegend()
        {
            return new XDDFChartLegend(chart);
        }

        public void deleteLegend()
        {
            if(chart.IsSetLegend())
            {
                chart.UnsetLegend();
            }
        }

        public XDDFManualLayout GetOrAddManualLayout()
        {
            return new XDDFManualLayout(chart.plotArea);
        }

        public void Plot<T, V>(XDDFChartData<T, V> data)
        {
            XSSFSheet sheet = GetSheet();
            foreach(XDDFChartData<T, V>.Series series in data.GetSeries())
            {
                series.Plot();
                FillSheet(sheet, series.GetCategoryData(), series.GetValuesData());
            }
        }

        public List<XDDFChartData<T, V>> GetChartSeries<T, V>()
        {
            List<XDDFChartData<T, V>> series = new List<XDDFChartData<T, V>>();
            CT_PlotArea plotArea = GetCTPlotArea();
            Dictionary<long, XDDFChartAxis> categories = GetCategoryAxes();
            Dictionary<long, XDDFValueAxis> values = GetValueAxes();

            for(int i = 0; i < plotArea.SizeOfBarChartArray(); i++)
            {
                CT_BarChart barChart = plotArea.GetBarChartArray(i);
                series.Add(new XDDFBarChartData<T, V>(barChart, categories, values));
            }

            for(int i = 0; i < plotArea.SizeOfLineChartArray(); i++)
            {
                CT_LineChart lineChart = plotArea.GetLineChartArray(i);
                series.Add(new XDDFLineChartData<T, V>(lineChart, categories, values));
            }

            for(int i = 0; i < plotArea.SizeOfPieChartArray(); i++)
            {
                CT_PieChart pieChart = plotArea.GetPieChartArray(i);
                series.Add(new XDDFPieChartData<T, V>(pieChart));
            }

            for(int i = 0; i < plotArea.SizeOfRadarChartArray(); i++)
            {
                CT_RadarChart radarChart = plotArea.GetRadarChartArray(i);
                series.Add(new XDDFRadarChartData<T, V>(radarChart, categories, values));
            }

            for(int i = 0; i < plotArea.SizeOfScatterChartArray(); i++)
            {
                CT_ScatterChart scatterChart = plotArea.GetScatterChartArray(i);
                series.Add(new XDDFScatterChartData<T, V>(scatterChart, categories, values));
            }

            // TODO repeat above code for all kind of charts
            return series;
        }

        private Dictionary<long, XDDFChartAxis> GetCategoryAxes()
        {
            CT_PlotArea plotArea = GetCTPlotArea();
            int sizeOfArray = plotArea.SizeOfCatAxArray();
            Dictionary<long, XDDFChartAxis> axes = new Dictionary<long, XDDFChartAxis>(sizeOfArray);
            for(int i = 0; i < sizeOfArray; i++)
            {
                CT_CatAx category = plotArea.GetCatAxArray(i);
                axes.Add(category.axId.val, new XDDFCategoryAxis(category));
            }
            return axes;
        }

        private Dictionary<long, XDDFValueAxis> GetValueAxes()
        {
            CT_PlotArea plotArea = GetCTPlotArea();
            int sizeOfArray = plotArea.SizeOfValAxArray();
            Dictionary<long, XDDFValueAxis> axes = new Dictionary<long, XDDFValueAxis>(sizeOfArray);
            for(int i = 0; i < sizeOfArray; i++)
            {
                CT_ValAx values = plotArea.GetValAxArray(i);
                axes.Add(values.axId.val, new XDDFValueAxis(values));
            }
            return axes;
        }

        public XDDFValueAxis CreateValueAxis(AxisPosition pos)
        {
            XDDFValueAxis valueAxis = new XDDFValueAxis(chart.plotArea, pos);
            if(axes.Count == 1)
            {
                XDDFChartAxis axis = axes[0];
                axis.crossAxis(valueAxis);
                valueAxis.crossAxis(axis);
            }
            axes.Add(valueAxis);
            return valueAxis;
        }

        public XDDFCategoryAxis CreateCategoryAxis(AxisPosition pos)
        {
            XDDFCategoryAxis categoryAxis = new XDDFCategoryAxis(chart.plotArea, pos);
            if(axes.Count == 1)
            {
                XDDFChartAxis axis = axes[0];
                axis.crossAxis(categoryAxis);
                categoryAxis.crossAxis(axis);
            }
            axes.Add(categoryAxis);
            return categoryAxis;
        }

        public XDDFDateAxis CreateDateAxis(AxisPosition pos)
        {
            XDDFDateAxis dateAxis = new XDDFDateAxis(chart.plotArea, pos);
            if(axes.Count == 1)
            {
                XDDFChartAxis axis = axes[0];
                axis.crossAxis(dateAxis);
                dateAxis.crossAxis(axis);
            }
            axes.Add(dateAxis);
            return dateAxis;
        }

        public XDDFChartData<T, V> CreateData<T, V>(ChartTypes type, 
            XDDFChartAxis category, XDDFValueAxis values)
        {
            Dictionary<long, XDDFChartAxis> categories = 
                new Dictionary<long, XDDFChartAxis>(){ { category.GetId(), category } } ;

            Dictionary<long, XDDFValueAxis> mapValues =
                new Dictionary<long, XDDFValueAxis>() { { values.GetId(), values } };

            CT_PlotArea plotArea = GetCTPlotArea();
            switch(type)
            {
                case ChartTypes.BAR:
                    return new XDDFBarChartData<T, V>(plotArea.AddNewBarChart(), categories, mapValues);
                case ChartTypes.LINE:
                    return new XDDFLineChartData<T, V>(plotArea.AddNewLineChart(), categories, mapValues);
                case ChartTypes.PIE:
                    return new XDDFPieChartData<T, V>(plotArea.AddNewPieChart());
                case ChartTypes.RADAR:
                    return new XDDFRadarChartData<T, V>(plotArea.AddNewRadarChart(), categories, mapValues);
                case ChartTypes.SCATTER:
                    return new XDDFScatterChartData<T, V>(plotArea.AddNewScatterChart(), categories, mapValues);
                default:
                    return null;
            }
        }

        public List<XDDFChartAxis> GetAxes()
        {
            if(axes.Count == 0 && HasAxes())
            {
                ParseAxes();
            }
            return axes;
        }

        private bool HasAxes()
        {
            CT_PlotArea ctPlotArea = chart.plotArea;
            int totalAxisCount = ctPlotArea.SizeOfValAxArray() + ctPlotArea.SizeOfCatAxArray() + ctPlotArea
            .SizeOfDateAxArray() + ctPlotArea.SizeOfSerAxArray();
            return totalAxisCount > 0;
        }

        private void ParseAxes()
        {
            foreach(CT_CatAx catAx in chart.plotArea.catAx)
            {
                axes.Add(new XDDFCategoryAxis(catAx));
            }
            foreach(CT_DateAx dateAx in chart.plotArea.dateAx)
            {
                axes.Add(new XDDFDateAxis(dateAx));
            }
            foreach(CT_SerAx serAx in chart.plotArea.serAx)
            {
                axes.Add(new XDDFSeriesAxis(serAx));
            }
            foreach(CT_ValAx valAx in chart.plotArea.valAx)
            {
                axes.Add(new XDDFValueAxis(valAx));
            }
        }

        /// <summary>
        /// Set value range (basic Axis Options)
        /// </summary>
        /// <param name="axisIndex">
        /// 0 - primary axis, 1 - secondary axis
        /// </param>
        /// <param name="minimum">
        /// minimum value; Double.NaN - automatic; null - no change
        /// </param>
        /// <param name="maximum">
        /// maximum value; Double.NaN - automatic; null - no change
        /// </param>
        /// <param name="majorUnit">
        /// major unit value; Double.NaN - automatic; null - no change
        /// </param>
        /// <param name="minorUnit">
        /// minor unit value; Double.NaN - automatic; null - no change
        /// </param>
        public void SetValueRange(int axisIndex,
            Double? minimum, Double? maximum,
            Double? majorUnit, Double? minorUnit)
        {
            XDDFChartAxis axis = GetAxes()[axisIndex];
            if(axis == null)
            {
                return;
            }
            if(minimum.HasValue)
            {
                axis.SetMinimum(minimum.Value);
            }
            if(maximum.HasValue)
            {
                axis.SetMaximum(maximum.Value);
            }
            if(majorUnit.HasValue)
            {
                axis.SetMajorUnit(majorUnit.Value);
            }
            if(minorUnit.HasValue)
            {
                axis.SetMinorUnit(minorUnit.Value);
            }
        }

        /// <summary>
        /// method to create relationship with embedded part for example writing xlsx
        /// file stream into output stream
        /// </summary>
        /// <param name="chartRelation">
        /// relationship object
        /// </param>
        /// <param name="chartFactory">
        /// ChartFactory object
        /// </param>
        /// <param name="chartIndex">
        /// index used to suffix on file
        /// </param>
        /// <returns>return relation part which used to write relation in .rels file
        /// and Get relation id
        /// </returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public PackageRelationship createRelationshipInChart(POIXMLRelation chartRelation, POIXMLFactory chartFactory,
            int chartIndex)
        {
            documentPart = CreateRelationship(chartRelation, chartFactory, chartIndex, true).DocumentPart;
            return this.AddRelation(null, chartRelation, documentPart).Relationship;
        }

        /// <summary>
        /// if embedded part was null then create new part
        /// </summary>
        /// <param name="chartRelation">
        /// chart relation object
        /// </param>
        /// <param name="chartWorkbookRelation">
        /// chart workbook relation object
        /// </param>
        /// <param name="chartFactory">
        /// factory object of POIXMLFactory (XWPFFactory/XSLFFactory)
        /// </param>
        /// <returns>return the new package part</returns>
        /// <exception cref="InvalidFormatException"></exception>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        private PackagePart createWorksheetPart(POIXMLRelation chartRelation, POIXMLRelation chartWorkbookRelation,
            POIXMLFactory chartFactory)
        {

            PackageRelationship xlsx = createRelationshipInChart(chartWorkbookRelation, chartFactory, chartIndex);
            this.SetExternalId(xlsx.Id);
            return GetTargetPart(xlsx);
        }

        /// <summary>
        /// this method write the XSSFWorkbook object data into embedded excel file
        /// </summary>
        /// <param name="workbook">
        /// XSSFworkbook object
        /// </param>
        /// <exception cref="IOException"></exception>
        /// <exception cref="InvalidFormatException"></exception>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public void SaveWorkbook(XSSFWorkbook workbook)
        {

            PackagePart worksheetPart = GetWorksheetPart();
            if(worksheetPart == null)
            {
                POIXMLRelation chartRelation = GetChartRelation();
                POIXMLRelation chartWorkbookRelation = GetChartWorkbookRelation();
                POIXMLFactory chartFactory = GetChartFactory();
                if(chartRelation != null && chartWorkbookRelation != null && chartFactory != null)
                {
                    worksheetPart = createWorksheetPart(chartRelation, chartWorkbookRelation, chartFactory);
                }
                else
                {
                    throw new InvalidFormatException("unable to determine chart relations");
                }
            }
            using(Stream xlsOut = worksheetPart.GetOutputStream())
            {
                SetWorksheetPartCommitted();
                workbook.Write(xlsOut);
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the chart relation in the implementing subclass.</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        protected abstract POIXMLRelation GetChartRelation();

        /// <summary>
        /// </summary>
        /// <returns>the chart workbook relation in the implementing subclass.</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        protected abstract POIXMLRelation GetChartWorkbookRelation();

        /// <summary>
        /// </summary>
        /// <returns>the chart factory in the implementing subclass.</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        protected abstract POIXMLFactory GetChartFactory();

        /// <summary>
        /// this method writes the data into sheet
        /// </summary>
        /// <param name="sheet">
        /// sheet of embedded excel
        /// </param>
        /// <param name="categoryData">
        /// category values
        /// </param>
        /// <param name="valuesData">
        /// data values
        /// </param>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        protected void FillSheet<T,V>(XSSFSheet sheet, 
            IXDDFDataSource<T> categoryData,
            IXDDFNumericalDataSource<V> valuesData)
        {
            int numOfPoints = categoryData.PointCount;
            for(int i = 0; i < numOfPoints; i++)
            {
                XSSFRow row = this.GetRow(sheet, i + 1); // first row is for title
                this.GetCell(row, categoryData.ColIndex).SetCellValue(categoryData.GetPointAt(i).ToString());
                this.GetCell(row, valuesData.ColIndex).SetCellValue(Convert.ToDouble(valuesData.GetPointAt(i)));
            }
        }

        /// <summary>
        /// this method return row on given index if row is null then create new row
        /// </summary>
        /// <param name="sheet">
        /// current sheet object
        /// </param>
        /// <param name="index">
        /// index of current row
        /// </param>
        /// <returns>this method return sheet row on given index</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        private XSSFRow GetRow(XSSFSheet sheet, int index)
        {
            if(sheet.GetRow(index) != null)
            {
                return sheet.GetRow(index) as XSSFRow;
            }
            else
            {
                return sheet.CreateRow(index) as XSSFRow;
            }
        }

        /// <summary>
        /// this method return cell on given index if cell is null then create new
        /// cell
        /// </summary>
        /// <param name="row">
        /// current row object
        /// </param>
        /// <param name="index">
        /// index of current cell
        /// </param>
        /// <returns>this method return sheet cell on given index</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        private XSSFCell GetCell(XSSFRow row, int index)
        {
            if(row.GetCell(index) != null)
            {
                return row.GetCell(index) as XSSFCell;
            }
            else
            {
                return row.CreateCell(index) as XSSFCell;
            }
        }

        /// <summary>
        /// import content from other chart to created chart
        /// </summary>
        /// <param name="other">
        /// chart object
        /// </param>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public void ImportContent(XDDFChart other)
        {
            this.chart.Set(other.chart);
        }

        /// <summary>
        /// save chart xml
        /// </summary>
        protected void Commit()
        {
            
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //xmlOptions.SetSaveSyntheticDocumentElement(
            //    new QName(CTChartSpace.type.Name.NamespaceURI, "chartSpace", "c"));

            if(workbook != null)
            {
                try
                {
                    SaveWorkbook(workbook);
                }
                catch(InvalidFormatException e)
                {
                    throw new POIXMLException(e);
                }
            }
            
            PackagePart part = GetPackagePart();
            using(Stream out1 = part.GetOutputStream())
            {
                new ChartSpaceDocument().Save(out1);
            }
        }

        /// <summary>
        /// Set sheet title in excel file
        /// </summary>
        /// <param name="title">
        /// title of sheet
        /// </param>
        /// <param name="column">
        /// column index
        /// </param>
        /// <returns>return cell reference</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public CellReference SetSheetTitle(string title, int column)
        {
            XSSFSheet sheet = GetSheet();
            XSSFRow row = this.GetRow(sheet, 0);
            XSSFCell cell = this.GetCell(row, column);
            cell.SetCellValue(title);
            this.updateSheetTable(sheet.GetTables()[0].GetCTTable(), title, column);
            return new CellReference(sheet.SheetName, 0, column, true, true);
        }

        /// <summary>
        /// this method update column header of sheet into table
        /// </summary>
        /// <param name="ctTable">
        /// xssf table object
        /// </param>
        /// <param name="title">
        /// title of column
        /// </param>
        /// <param name="index">
        /// index of column
        /// </param>
        private void updateSheetTable(OpenXmlFormats.Spreadsheet.CT_Table ctTable,
            string title, int index)
        {
            CT_TableColumns tableColumnList = ctTable.tableColumns;
            CT_TableColumn column = null;
            for(int i = 0; tableColumnList.count < index; i++)
            {
                column = tableColumnList.AddNewTableColumn();
                column.id = (uint) i;
            }
            column = tableColumnList.GetTableColumnArray(index);
            column.name = title;
        }

        /// <summary>
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public string FormatRange(CellRangeAddress range)
        {
            XSSFSheet sheet = GetSheet();
            return (sheet == null) ? null : range.FormatAsString(sheet.SheetName, true);
        }

        /// <summary>
        /// Get sheet object of embedded excel file
        /// </summary>
        /// <returns>excel sheet object</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        private XSSFSheet GetSheet()
        {
            XSSFSheet sheet = null;
            try
            {
                sheet = GetWorkbook().GetSheetAt(0) as XSSFSheet;
            }
            catch(InvalidFormatException ife)
            {
            }
            catch(IOException ioe)
            {
            }
            return sheet;
        }

        /// <summary>
        /// this method is used to Get worksheet part if call is from saveworkbook
        /// method then check isCommitted isCommitted variable shows that we are
        /// writing xssfworkbook object into output stream of embedded part
        /// </summary>
        /// <returns>returns the packagepart of embedded file</returns>
        /// <exception cref="InvalidFormatException"></exception>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        private PackagePart GetWorksheetPart()
        {

            foreach(RelationPart part in RelationParts)
            {
                if(POIXMLDocument.PACK_OBJECT_REL_TYPE.Equals(part.Relationship.RelationshipType))
                {
                    return GetTargetPart(part.Relationship);
                }
            }
            return null;
        }

        private void SetWorksheetPartCommitted()
        {
            foreach(RelationPart part in RelationParts)
            {
                if(POIXMLDocument.PACK_OBJECT_REL_TYPE.Equals(part.Relationship.RelationshipType))
                {
                    part.DocumentPart.Commited = true;
                    break;
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>returns the workbook object of embedded excel file</returns>
        /// <exception cref="IOException"></exception>
        /// <exception cref="InvalidFormatException"></exception>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public XSSFWorkbook GetWorkbook()
        {

            if(workbook == null)
            {
                try
                {
                    PackagePart worksheetPart = GetWorksheetPart();
                    if(worksheetPart == null)
                    {
                        workbook = new XSSFWorkbook();
                        workbook.CreateSheet();
                    }
                    else
                    {
                        workbook = new XSSFWorkbook(worksheetPart.GetInputStream());
                    }
                }
                catch(NotOfficeXmlFileException e)
                {
                    workbook = new XSSFWorkbook();
                    workbook.CreateSheet();
                }
            }
            return workbook;
        }

        /// <summary>
        /// while reading chart from template file then we need to parse and store
        /// embedded excel file in chart object show that we can modify value
        /// according to use
        /// </summary>
        /// <param name="workbook">
        /// workbook object which we read from chart embedded part
        /// </param>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public void SetWorkbook(XSSFWorkbook workbook)
        {
            this.workbook = workbook;
        }

        /// <summary>
        /// Set the relation id of embedded excel relation id into external data
        /// relation tag
        /// </summary>
        /// <param name="id">
        /// relation id of embedded excel relation id into external data
        /// relation tag
        /// </param>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        public void SetExternalId(string id)
        {
            GetCTChartSpace().AddNewExternalData().id = id;
        }

        /// <summary>
        /// </summary>
        /// <returns>method return chart index</returns>
        /// <remarks>
        /// @since POI 4.0.0
        /// </remarks>
        protected int GetChartIndex()
        {
            return chartIndex;
        }

        /// <summary>
        /// Set chart index which can be use for relation part
        /// </summary>
        /// <param name="chartIndex">
        /// chart index which can be use for relation part
        /// </param>
        public void SetChartIndex(int chartIndex)
        {
            this.chartIndex = chartIndex;
        }
    }
}

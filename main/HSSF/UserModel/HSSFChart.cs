/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Chart;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;


    public enum HSSFChartType : int
    {
        Area = 0x101A,
        Bar = 0x1017,
        Line = 0x1018,
        Pie = 0x1019,
        Scatter = 0x101B,
        Unknown = 0
    }
    /**
     * Has methods for construction of a chart object.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class HSSFChart
    {
        private HSSFSheet sheet;
        private ChartRecord chartRecord;

        private LegendRecord legendRecord;
        private AlRunsRecord chartTitleFormat;
        private SeriesTextRecord chartTitleText;
        private List<ValueRangeRecord> valueRanges = new List<ValueRangeRecord>();

        private HSSFChartType type = HSSFChartType.Unknown;

        private List<HSSFSeries> series = new List<HSSFSeries>();


        private HSSFChart(HSSFSheet sheet, ChartRecord chartRecord)
        {
            this.chartRecord = chartRecord;
            this.sheet = sheet;
        }

        /**
         * Creates a bar chart.  API needs some work. :)
         *
         * NOTE:  Does not yet work...  checking it in just so others
         * can take a look.
         */
        public void CreateBarChart(HSSFWorkbook workbook, HSSFSheet sheet)
        {

            List<RecordBase> records = new List<RecordBase>();
            records.Add(CreateMSDrawingObjectRecord());
            records.Add(CreateOBJRecord());
            records.Add(CreateBOFRecord());
            records.Add(new HeaderRecord(string.Empty));
            records.Add(new FooterRecord(string.Empty));
            records.Add(CreateHCenterRecord());
            records.Add(CreateVCenterRecord());
            records.Add(CreatePrintSetupRecord());
            // unknown 33   
            records.Add(CreateFontBasisRecord1());
            records.Add(CreateFontBasisRecord2());
            records.Add(new ProtectRecord(false));
            records.Add(CreateUnitsRecord());
            records.Add(CreateChartRecord(0, 0, 30434904, 19031616));
            records.Add(CreateBeginRecord());
            records.Add(CreateSCLRecord((short)1, (short)1));
            records.Add(CreatePlotGrowthRecord(65536, 65536));
            records.Add(CreateFrameRecord1());
            records.Add(CreateBeginRecord());
            records.Add(CreateLineFormatRecord(true));
            records.Add(CreateAreaFormatRecord1());
            records.Add(CreateEndRecord());
            records.Add(CreateSeriesRecord());
            records.Add(CreateBeginRecord());
            records.Add(CreateTitleLinkedDataRecord());
            records.Add(CreateValuesLinkedDataRecord());
            records.Add(CreateCategoriesLinkedDataRecord());
            records.Add(CreateDataFormatRecord());
            //		records.Add(CreateBeginRecord());
            // unknown
            //		records.Add(CreateEndRecord());
            records.Add(new SerToCrtRecord());
            records.Add(CreateEndRecord());
            records.Add(CreateSheetPropsRecord());
            records.Add(CreateDefaultTextRecord((short)TextFormatInfo.FontScaleNotSet));
            records.Add(CreateAllTextRecord());
            records.Add(CreateBeginRecord());
            // unknown
            records.Add(CreateFontIndexRecord(5));
            records.Add(CreateDirectLinkRecord());
            records.Add(CreateEndRecord());
            records.Add(CreateDefaultTextRecord((short)3)); // eek, undocumented text type
            records.Add(CreateUnknownTextRecord());
            records.Add(CreateBeginRecord());
            records.Add(CreateFontIndexRecord((short)6));
            records.Add(CreateDirectLinkRecord());
            records.Add(CreateEndRecord());

            records.Add(CreateAxisUsedRecord((short)1));
            CreateAxisRecords(records);

            records.Add(CreateEndRecord());
            records.Add(CreateDimensionsRecord());
            records.Add(CreateSeriesIndexRecord(2));
            records.Add(CreateSeriesIndexRecord(1));
            records.Add(CreateSeriesIndexRecord(3));
            records.Add(EOFRecord.instance);



            sheet.InsertChartRecords(records);
            workbook.InsertChartRecord();
        }

        /**
         * Returns all the charts for the given sheet.
         * 
         * NOTE: You won't be able to do very much with
         *  these charts yet, as this is very limited support
         */
        public static HSSFChart[] GetSheetCharts(HSSFSheet sheet)
        {
            List<HSSFChart> charts = new List<HSSFChart>();
            HSSFChart lastChart = null;
            HSSFSeries lastSeries = null;
            // Find records of interest
            IList records = sheet.Sheet.Records;
            foreach (RecordBase r in records)
            {

                if (r is ChartRecord)
                {
                    lastSeries = null;

                    lastChart = new HSSFChart(sheet, (ChartRecord)r);
                    charts.Add(lastChart);
                }
                else if (r is LegendRecord)
                {
                    lastChart.legendRecord = (LegendRecord)r;
                }
                else if (r is SeriesRecord)
                {
                    HSSFSeries series = new HSSFSeries((SeriesRecord)r);
                    lastChart.series.Add(series);
                    lastSeries = series;
                }
                else if (r is AlRunsRecord)
                {
                    lastChart.chartTitleFormat =
                        (AlRunsRecord)r;
                }
                else if (r is SeriesTextRecord)
                {
                    // Applies to a series, unless we've seen
                    //  a legend already
                    SeriesTextRecord str = (SeriesTextRecord)r;
                    if (lastChart.legendRecord == null &&
                            lastChart.series.Count > 0)
                    {
                        HSSFSeries series = (HSSFSeries)
                            lastChart.series[lastChart.series.Count - 1];
                        series.seriesTitleText = str;
                    }
                    else
                    {
                        lastChart.chartTitleText = str;
                    }
                }
                else if (r is BRAIRecord)
                {
                    BRAIRecord linkedDataRecord = (BRAIRecord)r;
                    if (lastSeries != null)
                    {
                        lastSeries.InsertData(linkedDataRecord);
                    }
                }
                else if (r is ValueRangeRecord)
                {
                    lastChart.valueRanges.Add((ValueRangeRecord)r);
                }
                else if (r is Record)
                {
                    if (lastChart != null)
                    {
                        Record record = (Record)r;
                        foreach (int type in Enum.GetValues(typeof(HSSFChartType)))
                        {
                            if (type == 0)
                            {
                                continue;
                            }
                            if (record.Sid == type)
                            {
                                lastChart.type = (HSSFChartType)type;
                                break;
                            }
                        }
                    }
                }
            }

            return (HSSFChart[])
                charts.ToArray();
        }

        /** Get the X offset of the chart */
        public int ChartX
        {
            get
            {
                return chartRecord.X;
            }
            set
            {
                chartRecord.X = value;
            }
        }
        /** Get the Y offset of the chart */
        public int ChartY
        {
            get { return chartRecord.Y; }
            set { chartRecord.Y = value; }
        }
        /** Get the width of the chart. {@link ChartRecord} */
        public int ChartWidth
        {
            get
            {
                return chartRecord.Width;
            }
            set
            {
                chartRecord.Width = value;
            }
        }
        /** Get the height of the chart. {@link ChartRecord} */
        public int ChartHeight
        {
            get
            {
                return chartRecord.Height;
            }
            set
            {
                chartRecord.Height = value;
            }
        }

        /**
         * Returns the series of the chart
         */
        public HSSFSeries[] Series
        {
            get
            {
                return (HSSFSeries[])
                    series.ToArray();
            }
        }

        /**
         * Returns the chart's title, if there is one,
         *  or null if not
         */
        public String ChartTitle
        {
            get
            {
                if (chartTitleText != null)
                {
                    return chartTitleText.Text;
                }
                return null;
            }
            set
            {
                if (chartTitleText != null)
                {
                    chartTitleText.Text = value;
                }
                else
                {
                    throw new InvalidOperationException("No chart title found to change");
                }
            }
        }

        /**
         * Set value range (basic Axis Options) 
         * @param axisIndex 0 - primary axis, 1 - secondary axis
         * @param minimum minimum value; Double.NaN - automatic; null - no change
         * @param maximum maximum value; Double.NaN - automatic; null - no change
         * @param majorUnit major unit value; Double.NaN - automatic; null - no change
         * @param minorUnit minor unit value; Double.NaN - automatic; null - no change
         */
        public void SetValueRange(int axisIndex, Double? minimum, Double? maximum, Double? majorUnit, Double? minorUnit)
        {
            ValueRangeRecord valueRange = (ValueRangeRecord)valueRanges[axisIndex];
            if (valueRange == null) return;
            if (minimum != null)
            {
                valueRange.IsAutomaticMinimum = Double.IsNaN((double)minimum);
                valueRange.MinimumAxisValue = (double)minimum;
            }
            if (maximum != null)
            {
                valueRange.IsAutomaticMaximum = Double.IsNaN((double)maximum);
                valueRange.MaximumAxisValue = (double)maximum;
            }
            if (majorUnit != null)
            {
                valueRange.IsAutomaticMajor = Double.IsNaN((double)majorUnit);
                valueRange.MajorIncrement = (double)majorUnit;
            }
            if (minorUnit != null)
            {
                valueRange.IsAutomaticMinor = Double.IsNaN((double)minorUnit);
                valueRange.MinorIncrement = (double)minorUnit;
            }
        }

        private SeriesIndexRecord CreateSeriesIndexRecord(int index)
        {
            SeriesIndexRecord r = new SeriesIndexRecord();
            r.Index = ((short)index);
            return r;
        }

        private DimensionsRecord CreateDimensionsRecord()
        {
            DimensionsRecord r = new DimensionsRecord();
            r.FirstRow = (0);
            r.LastRow = (31);
            r.FirstCol = ((short)0);
            r.LastCol = ((short)1);
            return r;
        }

        private HCenterRecord CreateHCenterRecord()
        {
            HCenterRecord r = new HCenterRecord();
            r.HCenter = (false);
            return r;
        }

        private VCenterRecord CreateVCenterRecord()
        {
            VCenterRecord r = new VCenterRecord();
            r.VCenter = (false);
            return r;
        }

        private PrintSetupRecord CreatePrintSetupRecord()
        {
            PrintSetupRecord r = new PrintSetupRecord();
            r.PaperSize = ((short)0);
            r.Scale = ((short)18);
            r.PageStart = ((short)1);
            r.FitWidth = ((short)1);
            r.FitHeight = ((short)1);
            r.LeftToRight = (false);
            r.Landscape = (false);
            r.ValidSettings = (true);
            r.NoColor = (false);
            r.Draft = (false);
            r.Notes = (false);
            r.NoOrientation = (false);
            r.UsePage = (false);
            r.HResolution = ((short)0);
            r.VResolution = ((short)0);
            r.HeaderMargin = (0.5);
            r.FooterMargin = (0.5);
            r.Copies = ((short)15); // what the ??
            return r;
        }

        private FbiRecord CreateFontBasisRecord1()
        {
            FbiRecord r = new FbiRecord();
            r.XBasis = ((short)9120);
            r.YBasis = ((short)5640);
            r.HeightBasis = ((short)200);
            r.Scale = ((short)0);
            r.IndexToFontTable = ((short)5);
            return r;
        }

        private FbiRecord CreateFontBasisRecord2()
        {
            FbiRecord r = CreateFontBasisRecord1();
            r.IndexToFontTable = ((short)6);
            return r;
        }

        private BOFRecord CreateBOFRecord()
        {
            BOFRecord r = new BOFRecord();
            r.Version = ((short)600);
            r.Type = BOFRecordType.Chart;
            r.Build = ((short)0x1CFE);
            r.BuildYear = ((short)1997);
            r.HistoryBitMask = (0x40C9);
            r.RequiredVersion = (106);
            return r;
        }

        private UnknownRecord CreateOBJRecord()
        {
            byte[] data = {
			(byte)0x15, (byte)0x00, (byte)0x12, (byte)0x00, (byte)0x05, (byte)0x00, (byte)0x02, (byte)0x00, 
            (byte)0x11, (byte)0x60, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0xB8, (byte)0x03,
			(byte)0x87, (byte)0x03, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, 
            (byte)0x00, (byte)0x00,
		};

            return new UnknownRecord((short)0x005D, data);
        }

        private UnknownRecord CreateMSDrawingObjectRecord()
        {
            // Since we haven't Created this object yet we'll just put in the raw
            // form for the moment.

            byte[] data = {
			    (byte)0x0F, (byte)0x00, (byte)0x02, (byte)0xF0, (byte)0xC0, (byte)0x00, (byte)0x00, (byte)0x00, 
                (byte)0x10, (byte)0x00, (byte)0x08, (byte)0xF0, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x02, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x02, (byte)0x04, (byte)0x00, (byte)0x00, 
                (byte)0x0F, (byte)0x00, (byte)0x03, (byte)0xF0, (byte)0xA8, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x0F, (byte)0x00, (byte)0x04, (byte)0xF0, (byte)0x28, (byte)0x00, (byte)0x00, (byte)0x00, 
                (byte)0x01, (byte)0x00, (byte)0x09, (byte)0xF0, (byte)0x10, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, 
                (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x02, (byte)0x00, (byte)0x0A, (byte)0xF0, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00, 
                (byte)0x00, (byte)0x04, (byte)0x00, (byte)0x00, (byte)0x05, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x0F, (byte)0x00, (byte)0x04, (byte)0xF0, (byte)0x70, (byte)0x00, (byte)0x00, (byte)0x00, 
                (byte)0x92, (byte)0x0C, (byte)0x0A, (byte)0xF0, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x02, (byte)0x04, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x0A, (byte)0x00, (byte)0x00, 
                (byte)0x93, (byte)0x00, (byte)0x0B, (byte)0xF0, (byte)0x36, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x7F, (byte)0x00, (byte)0x04, (byte)0x01, (byte)0x04, (byte)0x01, (byte)0xBF, (byte)0x00, 
                (byte)0x08, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x81, (byte)0x01, (byte)0x4E, (byte)0x00,
			    (byte)0x00, (byte)0x08, (byte)0x83, (byte)0x01, (byte)0x4D, (byte)0x00, (byte)0x00, (byte)0x08, 
                (byte)0xBF, (byte)0x01, (byte)0x10, (byte)0x00, (byte)0x11, (byte)0x00, (byte)0xC0, (byte)0x01,
			    (byte)0x4D, (byte)0x00, (byte)0x00, (byte)0x08, (byte)0xFF, (byte)0x01, (byte)0x08, (byte)0x00,
                (byte)0x08, (byte)0x00, (byte)0x3F, (byte)0x02, (byte)0x00, (byte)0x00, (byte)0x02, (byte)0x00,
			    (byte)0xBF, (byte)0x03, (byte)0x00, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00,
                (byte)0x10, (byte)0xF0, (byte)0x12, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
			    (byte)0x04, (byte)0x00, (byte)0xC0, (byte)0x02, (byte)0x0A, (byte)0x00, (byte)0xF4, (byte)0x00,
                (byte)0x0E, (byte)0x00, (byte)0x66, (byte)0x01, (byte)0x20, (byte)0x00, (byte)0xE9, (byte)0x00,
			    (byte)0x00, (byte)0x00, (byte)0x11, (byte)0xF0, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00
		    };

            return new UnknownRecord((short)0x00EC, data);
        }

        private void CreateAxisRecords(IList records)
        {
            records.Add(CreateAxisParentRecord());
            records.Add(CreateBeginRecord());
            records.Add(CreateAxisRecord(AxisRecord.AXIS_TYPE_CATEGORY_OR_X_AXIS));
            records.Add(CreateBeginRecord());
            records.Add(CreateCategorySeriesAxisRecord());
            records.Add(CreateAxisOptionsRecord());
            records.Add(CreateTickRecord1());
            records.Add(CreateEndRecord());
            records.Add(CreateAxisRecord(AxisRecord.AXIS_TYPE_VALUE_AXIS));
            records.Add(CreateBeginRecord());
            records.Add(CreateValueRangeRecord());
            records.Add(CreateTickRecord2());
            records.Add(CreateAxisLineFormatRecord(AxisLineRecord.AXIS_TYPE_MAJOR_GRID_LINE));
            records.Add(CreateLineFormatRecord(false));
            records.Add(CreateEndRecord());
            records.Add(CreatePlotAreaRecord());
            records.Add(CreateFrameRecord2());
            records.Add(CreateBeginRecord());
            records.Add(CreateLineFormatRecord2());
            records.Add(CreateAreaFormatRecord2());
            records.Add(CreateEndRecord());
            records.Add(CreateChartFormatRecord());
            records.Add(CreateBeginRecord());
            records.Add(CreateBarRecord());
            // unknown 1022
            records.Add(CreateLegendRecord());
            records.Add(CreateBeginRecord());
            // unknown 104f
            records.Add(CreateTextRecord());
            records.Add(CreateBeginRecord());
            // unknown 104f
            records.Add(CreateLinkedDataRecord());
            records.Add(CreateEndRecord());
            records.Add(CreateEndRecord());
            records.Add(CreateEndRecord());
            records.Add(CreateEndRecord());
        }

        private BRAIRecord CreateLinkedDataRecord()
        {
            BRAIRecord r = new BRAIRecord();
            r.LinkType = (BRAIRecord.LINK_TYPE_TITLE_OR_TEXT);
            r.ReferenceType = (BRAIRecord.REFERENCE_TYPE_DIRECT);
            r.IsCustomNumberFormat = (false);
            r.IndexNumberFmtRecord = ((short)0);
            r.FormulaOfLink = (null);
            return r;
        }

        private TextRecord CreateTextRecord()
        {
            TextRecord r = new TextRecord();
            r.HorizontalAlignment = (TextRecord.HORIZONTAL_ALIGNMENT_CENTER);
            r.VerticalAlignment = (TextRecord.VERTICAL_ALIGNMENT_CENTER);
            r.DisplayMode = ((short)1);
            r.RgbColor = (0x00000000);
            r.X = (-37);
            r.Y = (-60);
            r.Width = (0);
            r.Height = (0);
            r.IsAutoColor = (true);
            r.ShowKey = (false);
            r.ShowValue = (false);
            //r.IsVertical = (false);
            r.IsAutoGeneratedText = (true);
            r.IsGenerated = (true);
            r.IsAutoLabelDeleted = (false);
            r.IsAutoBackground = (true);
            //r.Rotation = ((short)0);
            r.ShowCategoryLabelAsPercentage = (false);
            r.ShowValueAsPercentage = (false);
            r.ShowBubbleSizes = (false);
            r.ShowLabel = (false);
            r.IndexOfColorValue = ((short)77);
            r.DataLabelPlacement = ((short)0);
            r.TextRotation = ((short)0);
            return r;
        }

        private LegendRecord CreateLegendRecord()
        {
            LegendRecord r = new LegendRecord();
            r.XAxisUpperLeft = (3542);
            r.YAxisUpperLeft = (1566);
            r.XSize = (437);
            r.YSize = (213);
            r.Type = (LegendRecord.TYPE_RIGHT);
            r.Spacing = (LegendRecord.SPACING_MEDIUM);
            r.IsAutoPosition = (true);
            r.IsAutoSeries = (true);
            r.IsAutoXPositioning = (true);
            r.IsAutoYPositioning = (true);
            r.IsVertical = (true);
            r.IsDataTable = (false);
            return r;
        }

        private BarRecord CreateBarRecord()
        {
            BarRecord r = new BarRecord();
            r.BarSpace = ((short)0);
            r.CategorySpace = ((short)150);
            r.IsHorizontal = (false);
            r.IsStacked = (false);
            r.IsDisplayAsPercentage = (false);
            r.IsShadow = (false);
            return r;
        }

        private ChartFormatRecord CreateChartFormatRecord()
        {
            ChartFormatRecord r = new ChartFormatRecord();
            r.XPosition = (0);
            r.YPosition = (0);
            r.Width = (0);
            r.Height = (0);
            r.VaryDisplayPattern = (false);
            return r;
        }

        private PlotAreaRecord CreatePlotAreaRecord()
        {
            PlotAreaRecord r = new PlotAreaRecord();
            return r;
        }

        private AxisLineRecord CreateAxisLineFormatRecord(short format)
        {
            AxisLineRecord r = new AxisLineRecord();
            r.AxisType = (format);
            return r;
        }

        private ValueRangeRecord CreateValueRangeRecord()
        {
            ValueRangeRecord r = new ValueRangeRecord();
            r.MinimumAxisValue = (0.0);
            r.MaximumAxisValue = (0.0);
            r.MajorIncrement = (0);
            r.MinorIncrement = (0);
            r.CategoryAxisCross = (0);
            r.IsAutomaticMinimum = (true);
            r.IsAutomaticMaximum = (true);
            r.IsAutomaticMajor = (true);
            r.IsAutomaticMinor = (true);
            r.IsAutomaticCategoryCrossing = (true);
            r.IsLogarithmicScale = (false);
            r.IsValuesInReverse = (false);
            r.IsCrossCategoryAxisAtMaximum = (false);
            r.IsReserved = (true);	// what's this do??
            return r;
        }

        private TickRecord CreateTickRecord1()
        {
            TickRecord r = new TickRecord();
            r.MajorTickType = ((byte)2);
            r.MinorTickType = ((byte)0);
            r.LabelPosition = ((byte)3);
            r.Background = ((byte)1);
            r.LabelColorRgb = (0);
            r.Zero1 = ((short)0);
            r.Zero2 = ((short)0);
            r.Zero3 = ((short)45);
            r.IsAutorotate = (true);
            r.IsAutoTextBackground = (true);
            r.Rotation = ((short)0);
            r.IsAutorotate = (true);
            r.TickColor = ((short)77);
            return r;
        }

        private TickRecord CreateTickRecord2()
        {
            TickRecord r = CreateTickRecord1();
            r.Zero3 = ((short)0);
            return r;
        }

        private AxcExtRecord CreateAxisOptionsRecord()
        {
            AxcExtRecord r = new AxcExtRecord();
            r.MinimumDate = ((short)-28644);
            r.MaximumDate = ((short)-28715);
            r.MajorInterval = ((short)2);
            r.MajorUnit = (DateUnit)0;
            r.MinorInterval = ((short)1);
            r.MinorUnit = (DateUnit)0;
            r.BaseUnit = (DateUnit)0;
            r.CrossDate = ((short)-28644);
            r.IsAutoMin = (true);
            r.IsAutoMax = (true);
            r.IsAutoMajor = (true);
            r.IsAutoMinor = (true);
            r.IsDateAxis = (true);
            r.IsAutoBase = (true);
            r.IsAutoCross = (true);
            r.IsAutoDate = (true);
            return r;
        }

        private CatSerRangeRecord CreateCategorySeriesAxisRecord()
        {
            CatSerRangeRecord r = new CatSerRangeRecord();
            r.CrossPoint = ((short)1);
            r.LabelInterval = ((short)1);
            r.MarkInterval = ((short)1);
            r.IsBetween = (true);
            r.IsMaxCross = (false);
            r.IsReverse = (false);
            return r;
        }

        private AxisRecord CreateAxisRecord(short axisType)
        {
            AxisRecord r = new AxisRecord();
            r.AxisType = (axisType);
            return r;
        }

        private AxisParentRecord CreateAxisParentRecord()
        {
            AxisParentRecord r = new AxisParentRecord();
            r.AxisType = (AxisParentRecord.AXIS_TYPE_MAIN);
            r.X = (479);
            r.Y = (221);
            r.Width = (2995);
            r.Height = (2902);
            return r;
        }

        private AxesUsedRecord CreateAxisUsedRecord(short numAxis)
        {
            AxesUsedRecord r = new AxesUsedRecord();
            r.NumAxis = (numAxis);
            return r;
        }

        private BRAIRecord CreateDirectLinkRecord()
        {
            BRAIRecord r = new BRAIRecord();
            r.LinkType = (BRAIRecord.LINK_TYPE_TITLE_OR_TEXT);
            r.ReferenceType = (BRAIRecord.REFERENCE_TYPE_DIRECT);
            r.IsCustomNumberFormat = (false);
            r.IndexNumberFmtRecord = ((short)0);
            r.FormulaOfLink = (null);
            return r;
        }

        private FontXRecord CreateFontIndexRecord(int index)
        {
            FontXRecord r = new FontXRecord();
            r.FontIndex = ((short)index);
            return r;
        }

        private TextRecord CreateAllTextRecord()
        {
            TextRecord r = new TextRecord();
            r.HorizontalAlignment = (TextRecord.HORIZONTAL_ALIGNMENT_CENTER);
            r.VerticalAlignment = (TextRecord.VERTICAL_ALIGNMENT_CENTER);
            r.DisplayMode = (TextRecord.DISPLAY_MODE_TRANSPARENT);
            r.RgbColor = (0);
            r.X = (-37);
            r.Y = (-60);
            r.Width = (0);
            r.Height = (0);
            r.IsAutoColor = (true);
            r.ShowKey = (false);
            r.ShowValue = (true);
            //r.IsVertical = (false);
            r.IsAutoGeneratedText = (true);
            r.IsGenerated = (true);
            r.IsAutoLabelDeleted = (false);
            r.IsAutoBackground = (true);
            //r.Rotation = ((short)0);
            r.ShowCategoryLabelAsPercentage = (false);
            r.ShowValueAsPercentage = (false);
            r.ShowBubbleSizes = (false);
            r.ShowLabel = (false);
            r.IndexOfColorValue = ((short)77);
            r.DataLabelPlacement = ((short)0);
            r.TextRotation = ((short)0);
            return r;
        }

        private TextRecord CreateUnknownTextRecord()
        {
            TextRecord r = new TextRecord();
            r.HorizontalAlignment = (TextRecord.HORIZONTAL_ALIGNMENT_CENTER);
            r.VerticalAlignment = (TextRecord.VERTICAL_ALIGNMENT_CENTER);
            r.DisplayMode = (TextRecord.DISPLAY_MODE_TRANSPARENT);
            r.RgbColor = (0);
            r.X = (-37);
            r.Y = (-60);
            r.Width = (0);
            r.Height = (0);
            r.IsAutoColor = (true);
            r.ShowKey = (false);
            r.ShowValue = (false);
            //r.IsVertical = (false);
            r.IsAutoGeneratedText = (true);
            r.IsGenerated = (true);
            r.IsAutoLabelDeleted = (false);
            r.IsAutoBackground = (true);
            //r.Rotation = ((short)0);
            r.ShowCategoryLabelAsPercentage = (false);
            r.ShowValueAsPercentage = (false);
            r.ShowBubbleSizes = (false);
            r.ShowLabel = (false);
            r.IndexOfColorValue = ((short)77);
            r.DataLabelPlacement = ((short)11088);
            r.TextRotation = ((short)0);
            return r;
        }

        private DefaultTextRecord CreateDefaultTextRecord(short categoryDataType)
        {
            DefaultTextRecord r = new DefaultTextRecord();
            r.FormatType = (TextFormatInfo)(categoryDataType);
            return r;
        }

        private ShtPropsRecord CreateSheetPropsRecord()
        {
            ShtPropsRecord r = new ShtPropsRecord();
            r.IsManSerAlloc = (false);
            r.IsPlotVisibleOnly = (true);
            r.IsNotSizeWithWindow = (false);
            r.IsManPlotArea = (true);
            r.IsAlwaysAutoPlotArea = (false);
            return r;
        }

        //private SeriesToChartGroupRecord CreateSeriesToChartGroupRecord()
        //{
        //    return new SeriesToChartGroupRecord();
        //}

        private DataFormatRecord CreateDataFormatRecord()
        {
            DataFormatRecord r = new DataFormatRecord();
            r.PointNumber = ((short)-1);
            r.SeriesIndex = ((short)0);
            r.SeriesNumber = ((short)0);
            r.UseExcel4Colors = (false);
            return r;
        }

        private BRAIRecord CreateCategoriesLinkedDataRecord()
        {
            BRAIRecord r = new BRAIRecord();
            r.LinkType = (BRAIRecord.LINK_TYPE_CATEGORIES);
            r.ReferenceType = (BRAIRecord.REFERENCE_TYPE_WORKSHEET);
            r.IsCustomNumberFormat = (false);
            r.IndexNumberFmtRecord = ((short)0);
            Area3DPtg p = new Area3DPtg(0, 31, 1, 1,
                    false, false, false, false, 0);
            r.FormulaOfLink = (new Ptg[] { p, });
            return r;
        }

        private BRAIRecord CreateValuesLinkedDataRecord()
        {
            BRAIRecord r = new BRAIRecord();
            r.LinkType = (BRAIRecord.LINK_TYPE_VALUES);
            r.ReferenceType = (BRAIRecord.REFERENCE_TYPE_WORKSHEET);
            r.IsCustomNumberFormat = (false);
            r.IndexNumberFmtRecord = ((short)0);
            Area3DPtg p = new Area3DPtg(0, 31, 0, 0,
                    false, false, false, false, 0);
            r.FormulaOfLink = (new Ptg[] { p, });
            return r;
        }

        private BRAIRecord CreateTitleLinkedDataRecord()
        {
            BRAIRecord r = new BRAIRecord();
            r.LinkType = (BRAIRecord.LINK_TYPE_TITLE_OR_TEXT);
            r.ReferenceType = (BRAIRecord.REFERENCE_TYPE_DIRECT);
            r.IsCustomNumberFormat = (false);
            r.IndexNumberFmtRecord = ((short)0);
            r.FormulaOfLink = (null);
            return r;
        }

        private SeriesRecord CreateSeriesRecord()
        {
            SeriesRecord r = new SeriesRecord();
            r.CategoryDataType = (SeriesRecord.CATEGORY_DATA_TYPE_NUMERIC);
            r.ValuesDataType = (SeriesRecord.VALUES_DATA_TYPE_NUMERIC);
            r.NumCategories = ((short)32);
            r.NumValues = ((short)31);
            r.BubbleSeriesType = (SeriesRecord.BUBBLE_SERIES_TYPE_NUMERIC);
            r.NumBubbleValues = ((short)0);
            return r;
        }

        private EndRecord CreateEndRecord()
        {
            return new EndRecord();
        }

        private AreaFormatRecord CreateAreaFormatRecord1()
        {
            AreaFormatRecord r = new AreaFormatRecord();
            r.ForegroundColor = (16777215);	 // RGB Color
            r.BackgroundColor = (0);			// RGB Color
            r.Pattern = ((short)1);			 // TODO: Add Pattern constants to record
            r.IsAutomatic = (true);
            r.IsInvert = (false);
            r.ForecolorIndex = ((short)78);
            r.BackcolorIndex = ((short)77);
            return r;
        }

        private AreaFormatRecord CreateAreaFormatRecord2()
        {
            AreaFormatRecord r = new AreaFormatRecord();
            r.ForegroundColor = (0x00c0c0c0);
            r.BackgroundColor = (0x00000000);
            r.Pattern = ((short)1);
            r.IsAutomatic = (false);
            r.IsInvert = (false);
            r.ForecolorIndex = ((short)22);
            r.BackcolorIndex = ((short)79);
            return r;
        }

        private LineFormatRecord CreateLineFormatRecord(bool drawTicks)
        {
            LineFormatRecord r = new LineFormatRecord();
            r.LineColor = (0);
            r.LinePattern = (LineFormatRecord.LINE_PATTERN_SOLID);
            r.Weight = ((short)-1);
            r.IsAuto = (true);
            r.IsDrawTicks = (drawTicks);
            r.ColourPaletteIndex = ((short)77);  // what colour is this?
            return r;
        }

        private LineFormatRecord CreateLineFormatRecord2()
        {
            LineFormatRecord r = new LineFormatRecord();
            r.LineColor = (0x00808080);
            r.LinePattern = ((short)0);
            r.Weight = ((short)0);
            r.IsAuto = (false);
            r.IsDrawTicks = (false);
            r.IsUnknown = (false);
            r.ColourPaletteIndex = ((short)23);
            return r;
        }

        private FrameRecord CreateFrameRecord1()
        {
            FrameRecord r = new FrameRecord();
            r.BorderType = (FrameRecord.BORDER_TYPE_REGULAR);
            r.IsAutoSize = (false);
            r.IsAutoPosition = (true);
            return r;
        }

        private FrameRecord CreateFrameRecord2()
        {
            FrameRecord r = new FrameRecord();
            r.BorderType = (FrameRecord.BORDER_TYPE_REGULAR);
            r.IsAutoSize = (true);
            r.IsAutoPosition = (true);
            return r;
        }

        private PlotGrowthRecord CreatePlotGrowthRecord(int horizScale, int vertScale)
        {
            PlotGrowthRecord r = new PlotGrowthRecord();
            r.HorizontalScale = (horizScale);
            r.VerticalScale = (vertScale);
            return r;
        }

        private SCLRecord CreateSCLRecord(short numerator, short denominator)
        {
            SCLRecord r = new SCLRecord();
            r.Denominator = (denominator);
            r.Numerator = (numerator);
            return r;
        }

        private BeginRecord CreateBeginRecord()
        {
            return new BeginRecord();
        }

        private ChartRecord CreateChartRecord(int x, int y, int width, int height)
        {
            ChartRecord r = new ChartRecord();
            r.X = (x);
            r.Y = (y);
            r.Width = (width);
            r.Height = (height);
            return r;
        }

        private UnitsRecord CreateUnitsRecord()
        {
            UnitsRecord r = new UnitsRecord();
            r.Units = ((short)0);
            return r;
        }


        /**
         * A series in a chart
         */
        public class HSSFSeries
        {
            internal SeriesRecord series;
            internal SeriesTextRecord seriesTitleText;
            private BRAIRecord dataName;
            private BRAIRecord dataValues;
            private BRAIRecord dataCategoryLabels;
            private BRAIRecord dataSecondaryCategoryLabels;

            /* package */
            public HSSFSeries(SeriesRecord series)
            {
                this.series = series;
            }

            internal void InsertData(BRAIRecord data)
            {
                switch (data.LinkType)
                {
                    case 0: dataName = data;
                        break;
                    case 1: dataValues = data;
                        break;
                    case 2: dataCategoryLabels = data;
                        break;
                    case 3: dataSecondaryCategoryLabels = data;
                        break;
                }
            }

            internal void SetSeriesTitleText(SeriesTextRecord seriesTitleText)
            {
                this.seriesTitleText = seriesTitleText;
            }

            public short NumValues
            {
                get
                {
                    return series.NumValues;
                }
            }
            /**
             * See {@link SeriesRecord}
             */
            public short ValueType
            {
                get
                {
                    return series.ValuesDataType;
                }
            }

            /**
             * Returns the series' title, if there is one,
             *  or null if not
             */
            public String SeriesTitle
            {
                get
                {
                    if (seriesTitleText != null)
                    {
                        return seriesTitleText.Text;
                    }
                    return null;
                }
                set
                {
                    if (seriesTitleText != null)
                    {
                        seriesTitleText.Text = value;
                    }
                    else
                    {
                        throw new InvalidOperationException("No series title found to Change");
                    }
                }
            }

            /**
             * @return record with data names
             */
            public BRAIRecord GetDataName()
            {
                return dataName;
            }

            /**
             * @return record with data values
             */
            public BRAIRecord GetDataValues()
            {
                return dataValues;
            }

            /**
             * @return record with data category labels
             */
            public BRAIRecord GetDataCategoryLabels()
            {
                return dataCategoryLabels;
            }

            /**
             * @return record with data secondary category labels
             */
            public BRAIRecord GetDataSecondaryCategoryLabels()
            {
                return dataSecondaryCategoryLabels;
            }

            /**
             * @return record with series
             */
            public SeriesRecord GetSeries()
            {
                return series;
            }

            private CellRangeAddressBase GetCellRange(BRAIRecord linkedDataRecord)
            {
                if (linkedDataRecord == null)
                {
                    return null;
                }

                int firstRow = 0;
                int lastRow = 0;
                int firstCol = 0;
                int lastCol = 0;

                foreach (Ptg ptg in linkedDataRecord.FormulaOfLink)
                {
                    if (ptg is AreaPtgBase)
                    {
                        AreaPtgBase areaPtg = (AreaPtgBase)ptg;

                        firstRow = areaPtg.FirstRow;
                        lastRow = areaPtg.LastRow;

                        firstCol = areaPtg.FirstColumn;
                        lastCol = areaPtg.LastColumn;
                    }
                }

                return new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            }

            public CellRangeAddressBase GetValuesCellRange()
            {
                return GetCellRange(dataValues);
            }

            public CellRangeAddressBase GetCategoryLabelsCellRange()
            {
                return GetCellRange(dataCategoryLabels);
            }

            private int SetVerticalCellRange(BRAIRecord linkedDataRecord,
                                                 CellRangeAddressBase range)
            {
                if (linkedDataRecord == null)
                {
                    throw new ArgumentNullException("linkedDataRecord should not be null"); ;
                }

                List<Ptg> ptgList = new List<Ptg>();

                int rowCount = (range.LastRow - range.FirstRow) + 1;
                int colCount = (range.LastColumn - range.FirstColumn) + 1;

                foreach (Ptg ptg in linkedDataRecord.FormulaOfLink)
                {
                    if (ptg is AreaPtgBase)
                    {
                        AreaPtgBase areaPtg = (AreaPtgBase)ptg;

                        areaPtg.FirstRow = range.FirstRow;
                        areaPtg.LastRow = range.LastRow;

                        areaPtg.FirstColumn = range.FirstColumn;
                        areaPtg.LastColumn = range.LastColumn;
                        ptgList.Add(areaPtg);
                    }
                }

                linkedDataRecord.FormulaOfLink = (ptgList.ToArray());

                return rowCount * colCount;
            }

            public void SetValuesCellRange(CellRangeAddressBase range)
            {
                int count = SetVerticalCellRange(dataValues, range);

                series.NumValues = (short)count;
            }

            public void SetCategoryLabelsCellRange(CellRangeAddressBase range)
            {
                int count = SetVerticalCellRange(dataCategoryLabels, range);

                series.NumCategories = (short)count;
            }
        }

        public HSSFSeries CreateSeries()
        {
            List<RecordBase> seriesTemplate = new List<RecordBase>();
            bool seriesTemplateFilled = false;

            int idx = 0;
            int deep = 0;
            int chartRecordIdx = -1;
            int chartDeep = -1;
            int lastSeriesDeep = -1;
            int endSeriesRecordIdx = -1;
            int seriesIdx = 0;
            IList records = sheet.Sheet.Records;

            /* store first series as template and find last series index */
            foreach (RecordBase record in records)
            {

                idx++;

                if (record is BeginRecord)
                {
                    deep++;
                }
                else if (record is EndRecord)
                {
                    deep--;

                    if (lastSeriesDeep == deep)
                    {
                        lastSeriesDeep = -1;
                        endSeriesRecordIdx = idx;
                        if (!seriesTemplateFilled)
                        {
                            seriesTemplate.Add(record);
                            seriesTemplateFilled = true;
                        }
                    }

                    if (chartDeep == deep)
                    {
                        break;
                    }
                }

                if (record is ChartRecord)
                {
                    if (record == chartRecord)
                    {
                        chartRecordIdx = idx;
                        chartDeep = deep;
                    }
                }
                else if (record is SeriesRecord)
                {
                    if (chartRecordIdx != -1)
                    {
                        seriesIdx++;
                        lastSeriesDeep = deep;
                    }
                }

                if (lastSeriesDeep != -1 && !seriesTemplateFilled)
                {
                    seriesTemplate.Add(record);
                }
            }

            /* check if a series was found */
            if (endSeriesRecordIdx == -1)
            {
                return null;
            }

            /* next index in the records list where the new series can be inserted */
            idx = endSeriesRecordIdx + 1;

            HSSFSeries newSeries = null;

            /* duplicate record of the template series */
            List<RecordBase> ClonedRecords = new List<RecordBase>();
            foreach (RecordBase record in seriesTemplate)
            {

                Record newRecord = null;

                if (record is BeginRecord)
                {
                    newRecord = new BeginRecord();
                }
                else if (record is EndRecord)
                {
                    newRecord = new EndRecord();
                }
                else if (record is SeriesRecord)
                {
                    SeriesRecord seriesRecord = (SeriesRecord)((SeriesRecord)record).Clone();
                    newSeries = new HSSFSeries(seriesRecord);
                    newRecord = seriesRecord;
                }
                else if (record is BRAIRecord)
                {
                    BRAIRecord linkedDataRecord = (BRAIRecord)((BRAIRecord)record).Clone();
                    if (newSeries != null)
                    {
                        newSeries.InsertData(linkedDataRecord);
                    }
                    newRecord = linkedDataRecord;
                }
                else if (record is DataFormatRecord)
                {
                    DataFormatRecord dataFormatRecord = (DataFormatRecord)((DataFormatRecord)record).Clone();

                    dataFormatRecord.SeriesIndex = ((short)seriesIdx);
                    dataFormatRecord.SeriesNumber = ((short)seriesIdx);

                    newRecord = dataFormatRecord;
                }
                else if (record is SeriesTextRecord)
                {
                    SeriesTextRecord seriesTextRecord = (SeriesTextRecord)((SeriesTextRecord)record).Clone();
                    if (newSeries != null)
                    {
                        newSeries.SetSeriesTitleText(seriesTextRecord);
                    }
                    newRecord = seriesTextRecord;
                }
                else if (record is Record)
                {
                    newRecord = (Record)((Record)record).Clone();
                }

                if (newRecord != null)
                {
                    ClonedRecords.Add(newRecord);
                }
            }

            /* check if a user model series object was Created */
            if (newSeries == null)
            {
                return null;
            }

            /* transfer series to record list */
            foreach (RecordBase record in ClonedRecords)
            {
                records.Insert(idx++, record);
            }

            return newSeries;
        }

        public bool RemoveSeries(HSSFSeries series)
        {
            int idx = 0;
            int deep = 0;
            int chartDeep = -1;
            int lastSeriesDeep = -1;
            int seriesIdx = -1;
            bool RemoveSeries = false;
            bool chartEntered = false;
            bool result = false;
            IList records = sheet.Sheet.Records;

            /* store first series as template and find last series index */

            IEnumerator iter = records.GetEnumerator();
            while (iter.MoveNext())
            {
                RecordBase record = (RecordBase)iter.Current;
                idx++;

                if (record is BeginRecord)
                {
                    deep++;
                }
                else if (record is EndRecord)
                {
                    deep--;

                    if (lastSeriesDeep == deep)
                    {
                        lastSeriesDeep = -1;

                        if (RemoveSeries)
                        {
                            RemoveSeries = false;
                            result = true;
                            records.Remove(record);
                        }
                    }

                    if (chartDeep == deep)
                    {
                        break;
                    }
                }

                if (record is ChartRecord)
                {
                    if (record == chartRecord)
                    {
                        chartDeep = deep;
                        chartEntered = true;
                    }
                }
                else if (record is SeriesRecord)
                {
                    if (chartEntered)
                    {
                        if (series.series == record)
                        {
                            lastSeriesDeep = deep;
                            RemoveSeries = true;
                        }
                        else
                        {
                            seriesIdx++;
                        }
                    }
                }
                else if (record is DataFormatRecord)
                {
                    if (chartEntered && !RemoveSeries)
                    {
                        DataFormatRecord dataFormatRecord = (DataFormatRecord)record;
                        dataFormatRecord.SeriesIndex = ((short)seriesIdx);
                        dataFormatRecord.SeriesNumber = ((short)seriesIdx);
                    }
                }

                if (RemoveSeries)
                {
                    records.Remove(record);
                }
            }

            return result;
        }

        public HSSFChartType Type
        {
            get
            {
                return type;
            }
        }
    }
}
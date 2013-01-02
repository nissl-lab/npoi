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

namespace NPOI.HSSF.UserModel.ScratchPad
{
using NPOI.HSSF.Record.Chart;
using NPOI.HSSF.Record;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Util;
using System.Collections.Generic;
    using System;

/**
 * Has methods for construction of a chart object.
 *
 * @author Glen Stampoultzis (glens at apache.org)
 */
public  class HSSFChart {
	private HSSFSheet sheet;
	private ChartRecord chartRecord;

	private LegendRecord legendRecord;
	private ChartTitleFormatRecord chartTitleFormat;
	private SeriesTextRecord chartTitleText;
	private List<ValueRangeRecord> valueRanges = new List<ValueRangeRecord>(); 
	
	private HSSFChartType type = HSSFChartType.Unknown;

    private List<ScratchPad.HSSFChart.HSSFSeries> series = new List<ScratchPad.HSSFChart.HSSFSeries>();

	private HSSFChart(HSSFSheet sheet, ChartRecord chartRecord) {
		this.chartRecord = chartRecord;
		this.sheet = sheet;
	}

	/**
	 * Creates a bar chart.  API needs some work. :)
	 * <p>
	 * NOTE:  Does not yet work...  checking it in just so others
	 * can take a look.
	 */
	public void CreateBarChart( HSSFWorkbook workbook, HSSFSheet sheet )
	{

		List<Record> records = new List<Record>();
		records.Add( CreateMSDrawingObjectRecord() );
		records.Add( CreateOBJRecord() );
		records.Add( CreateBOFRecord() );
		records.Add(new HeaderRecord(""));
		records.Add(new FooterRecord(""));
		records.Add( CreateHCenterRecord() );
		records.Add( CreateVCenterRecord() );
		records.Add( CreatePrintSetupRecord() );
		// unknown 33
		records.Add( CreateFontBasisRecord1() );
		records.Add( CreateFontBasisRecord2() );
		records.Add(new ProtectRecord(false));
		records.Add( CreateUnitsRecord() );
		records.Add( CreateChartRecord( 0, 0, 30434904, 19031616 ) );
		records.Add( CreateBeginRecord() );
		records.Add( CreateSCLRecord( (short) 1, (short) 1 ) );
		records.Add( CreatePlotGrowthRecord( 65536, 65536 ) );
		records.Add( CreateFrameRecord1() );
		records.Add( CreateBeginRecord() );
		records.Add( CreateLineFormatRecord(true) );
		records.Add( CreateAreaFormatRecord1() );
		records.Add( CreateEndRecord() );
		records.Add( CreateSeriesRecord() );
		records.Add( CreateBeginRecord() );
		records.Add( CreateTitleLinkedDataRecord() );
		records.Add( CreateValuesLinkedDataRecord() );
		records.Add( CreateCategoriesLinkedDataRecord() );
		records.Add( CreateDataFormatRecord() );
		//		records.add(createBeginRecord());
		// unknown
		//		records.add(createEndRecord());
		records.Add( CreateSeriesToChartGroupRecord() );
		records.Add( CreateEndRecord() );
		records.Add( CreateSheetPropsRecord() );
		records.Add( CreateDefaultTextRecord( DefaultDataLabelTextPropertiesRecord.CATEGORY_DATA_TYPE_ALL_TEXT_CHARACTERISTIC ) );
		records.Add( CreateAllTextRecord() );
		records.Add( CreateBeginRecord() );
		// unknown
		records.Add( CreateFontIndexRecord( 5 ) );
		records.Add( CreateDirectLinkRecord() );
		records.Add( CreateEndRecord() );
		records.Add( CreateDefaultTextRecord( (short) 3 ) ); // eek, undocumented text type
		records.Add( CreateUnknownTextRecord() );
		records.Add( CreateBeginRecord() );
		records.Add( CreateFontIndexRecord( (short) 6 ) );
		records.Add( CreateDirectLinkRecord() );
		records.Add( CreateEndRecord() );

		records.Add( CreateAxisUsedRecord( (short) 1 ) );
		CreateAxisRecords( records );

		records.Add( CreateEndRecord() );
		records.Add( CreateDimensionsRecord() );
		records.Add( CreateSeriesIndexRecord(2) );
		records.Add( CreateSeriesIndexRecord(1) );
		records.Add( CreateSeriesIndexRecord(3) );
		records.Add(EOFRecord.instance);



		sheet.InsertChartRecords( records );
		workbook.InsertChartRecord();
	}

	/**
	 * Returns all the charts for the given sheet.
	 *
	 * NOTE: You won't be able to do very much with
	 *  these charts yet, as this is very limited support
	 */
	public static HSSFChart[] GetSheetCharts(HSSFSheet sheet) {
		List<HSSFChart> charts = new List<HSSFChart>();
		HSSFChart lastChart = null;
		HSSFSeries lastSeries = null;
		// Find records of interest
		List<RecordBase> records = sheet.GetSheet().GetRecords();
		foreach(RecordBase r in records) {

			if(r is ChartRecord) {
				lastSeries = null;
				
				lastChart = new HSSFChart(sheet,(ChartRecord)r);
				charts.Add(lastChart);
			} else if(r is LegendRecord) {
				lastChart.legendRecord = (LegendRecord)r;
			} else if(r is SeriesRecord) {
				HSSFSeries series = lastChart.new HSSFSeries( (SeriesRecord)r );
				lastChart.series.Add(series);
				lastSeries = series;
			} else if(r is ChartTitleFormatRecord) {
				lastChart.chartTitleFormat =
					(ChartTitleFormatRecord)r;
			} else if(r is SeriesTextRecord) {
				// Applies to a series, unless we've seen
				//  a legend already
				SeriesTextRecord str = (SeriesTextRecord)r;
				if(lastChart.legendRecord == null &&
						lastChart.series.Size() > 0) {
					HSSFSeries series = (HSSFSeries)
						lastChart.series.Get(lastChart.series.Size()-1);
					series.seriesTitleText = str;
				} else {
					lastChart.chartTitleText = str;
				}
			} else if (r is LinkedDataRecord) {
				LinkedDataRecord linkedDataRecord = (LinkedDataRecord) r;
				if (lastSeries != null) {
					lastSeries.InsertData(linkedDataRecord);
				}
			} else if(r is ValueRangeRecord){
				lastChart.valueRanges.Add((ValueRangeRecord)r);
			} else if (r is Record) {
				if (lastChart != null)
				{
					Record record = (Record) r;
					for (HSSFChartType type : HSSFChartType.Values()) {
						if (type == HSSFChartType.Unknown)
						{
							continue;
						}
						if (record.GetSid() == type.GetSid()) {
							lastChart.type = type ;
							break;
						}
					}
				}
			}
		}

		return (HSSFChart[])
			charts.ToArray( new HSSFChart[charts.Size()] );
	}

	/** Get the X offset of the chart */
	public int GetChartX() { return chartRecord.GetX(); }
	/** Get the Y offset of the chart */
	public int GetChartY() { return chartRecord.GetY(); }
	/** Get the width of the chart. {@link ChartRecord} */
	public int GetChartWidth() { return chartRecord.GetWidth(); }
	/** Get the height of the chart. {@link ChartRecord} */
	public int GetChartHeight() { return chartRecord.GetHeight(); }

	/** Sets the X offset of the chart */
	public void SetChartX(int x) { chartRecord.SetX(x); }
	/** Sets the Y offset of the chart */
	public void SetChartY(int y) { chartRecord.SetY(y); }
	/** Sets the width of the chart. {@link ChartRecord} */
	public void SetChartWidth(int width) { chartRecord.SetWidth(width); }
	/** Sets the height of the chart. {@link ChartRecord} */
	public void SetChartHeight(int height) { chartRecord.SetHeight(height); }

	/**
	 * Returns the series of the chart
	 */
	public HSSFSeries[] GetSeries() {
		return (HSSFSeries[])
			series.ToArray(new HSSFSeries[series.Size()]);
	}

	/**
	 * Returns the chart's title, if there is one,
	 *  or null if not
	 */
	public String GetChartTitle() {
		if(chartTitleText != null) {
			return chartTitleText.GetText();
		}
		return null;
	}

	/**
	 * Changes the chart's title, but only if there
	 *  was one already.
	 * TODO - add in the records if not
	 */
	public void SetChartTitle(String title) {
		if(chartTitleText != null) {
			chartTitleText.SetText(title);
		} else {
			throw new IllegalStateException("No chart title found to change");
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
	public void SetValueRange( int axisIndex, Double minimum, Double maximum, Double majorUnit, Double minorUnit){
		ValueRangeRecord valueRange = (ValueRangeRecord)valueRanges[ axisIndex ];
		if( valueRange == null ) return;
		if( minimum != null ){
			valueRange.AutomaticMinimum=minimum.IsNaN();
			valueRange.MinimumAxisValue=minimum;
		}
		if( maximum != null ){
			valueRange.SetAutomaticMaximum(maximum.IsNaN());
			valueRange.SetMaximumAxisValue(maximum);
		}
		if( majorUnit != null ){
			valueRange.SetAutomaticMajor(majorUnit.IsNaN());
			valueRange.SetMajorIncrement(majorUnit);
		}
		if( minorUnit != null ){
			valueRange.SetAutomaticMinor(minorUnit.IsNaN());
			valueRange.SetMinorIncrement(minorUnit);
		}
	}

	private SeriesIndexRecord CreateSeriesIndexRecord( int index )
	{
		SeriesIndexRecord r = new SeriesIndexRecord();
		r.SetIndex((short)index);
		return r;
	}

	private DimensionsRecord CreateDimensionsRecord()
	{
		DimensionsRecord r = new DimensionsRecord();
		r.SetFirstRow(0);
		r.SetLastRow(31);
		r.SetFirstCol((short)0);
		r.SetLastCol((short)1);
		return r;
	}

	private HCenterRecord CreateHCenterRecord()
	{
		HCenterRecord r = new HCenterRecord();
		r.SetHCenter(false);
		return r;
	}

	private VCenterRecord CreateVCenterRecord()
	{
		VCenterRecord r = new VCenterRecord();
		r.SetVCenter(false);
		return r;
	}

	private PrintSetupRecord CreatePrintSetupRecord()
	{
		PrintSetupRecord r = new PrintSetupRecord();
		r.PaperSize=((short)0);
		r.Scale=((short)18);
		r.PageStart=((short)1);
		r.FitWidth=((short)1);
		r.FitHeight=((short)1);
		r.LeftToRight=(false);
		r.Landscape=(false);
		r.ValidSettings=(true);
		r.NoColor=(false);
		r.Draft=(false);
		r.Notes=(false);
		r.NoOrientation=(false);
		r.UsePage=(false);
		r.HResolution=((short)0);
		r.VResolution=((short)0);
		r.HeaderMargin=(0.5);
		r.FooterMargin=(0.5);
		r.Copies=((short)15); // what the ??
		return r;
	}

	private FontBasisRecord CreateFontBasisRecord1()
	{
		FontBasisRecord r = new FontBasisRecord();
		r.XBasis=((short)9120);
		r.YBasis=((short)5640);
		r.HeightBasis=((short)200);
		r.Scale=((short)0);
		r.IndexToFontTable=((short)5);
		return r;
	}

	private FontBasisRecord CreateFontBasisRecord2()
	{
		FontBasisRecord r = CreateFontBasisRecord1();
		r.IndexToFontTable=((short)6);
		return r;
	}

	private BOFRecord CreateBOFRecord()
	{
		BOFRecord r = new BOFRecord();
		r.Version=((short)600);
		r.Type=((short)20);
		r.Build=((short)0x1CFE);
		r.BuildYear=((short)1997);
		r.HistoryBitMask=(0x40C9);
		r.RequiredVersion=(106);
		return r;
	}

	private UnknownRecord CreateOBJRecord()
	{
		byte[] data = {
			(byte) 0x15, (byte) 0x00, (byte) 0x12, (byte) 0x00, (byte) 0x05, (byte) 0x00, (byte) 0x02, (byte) 0x00, (byte) 0x11, (byte) 0x60, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0xB8, (byte) 0x03,
			(byte) 0x87, (byte) 0x03, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
		};

		return new UnknownRecord( (short) 0x005D, data );
	}

	private UnknownRecord CreateMSDrawingObjectRecord()
	{
		// Since we haven't created this object yet we'll just put in the raw
		// form for the moment.

		byte[] data = {
			(byte)0x0F, (byte)0x00, (byte)0x02, (byte)0xF0, (byte)0xC0, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x10, (byte)0x00, (byte)0x08, (byte)0xF0, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x02, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x02, (byte)0x04, (byte)0x00, (byte)0x00, (byte)0x0F, (byte)0x00, (byte)0x03, (byte)0xF0, (byte)0xA8, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x0F, (byte)0x00, (byte)0x04, (byte)0xF0, (byte)0x28, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x01, (byte)0x00, (byte)0x09, (byte)0xF0, (byte)0x10, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x02, (byte)0x00, (byte)0x0A, (byte)0xF0, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x04, (byte)0x00, (byte)0x00, (byte)0x05, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x0F, (byte)0x00, (byte)0x04, (byte)0xF0, (byte)0x70, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x92, (byte)0x0C, (byte)0x0A, (byte)0xF0, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x02, (byte)0x04, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x0A, (byte)0x00, (byte)0x00, (byte)0x93, (byte)0x00, (byte)0x0B, (byte)0xF0, (byte)0x36, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x7F, (byte)0x00, (byte)0x04, (byte)0x01, (byte)0x04, (byte)0x01, (byte)0xBF, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x81, (byte)0x01, (byte)0x4E, (byte)0x00,
			(byte)0x00, (byte)0x08, (byte)0x83, (byte)0x01, (byte)0x4D, (byte)0x00, (byte)0x00, (byte)0x08, (byte)0xBF, (byte)0x01, (byte)0x10, (byte)0x00, (byte)0x11, (byte)0x00, (byte)0xC0, (byte)0x01,
			(byte)0x4D, (byte)0x00, (byte)0x00, (byte)0x08, (byte)0xFF, (byte)0x01, (byte)0x08, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x3F, (byte)0x02, (byte)0x00, (byte)0x00, (byte)0x02, (byte)0x00,
			(byte)0xBF, (byte)0x03, (byte)0x00, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x10, (byte)0xF0, (byte)0x12, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
			(byte)0x04, (byte)0x00, (byte)0xC0, (byte)0x02, (byte)0x0A, (byte)0x00, (byte)0xF4, (byte)0x00, (byte)0x0E, (byte)0x00, (byte)0x66, (byte)0x01, (byte)0x20, (byte)0x00, (byte)0xE9, (byte)0x00,
			(byte)0x00, (byte)0x00, (byte)0x11, (byte)0xF0, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00
		};

		return new UnknownRecord((short)0x00EC, data);
	}

	private void CreateAxisRecords( List<Record> records )
	{
		records.Add( CreateAxisParentRecord() );
		records.Add( CreateBeginRecord() );
		records.Add( CreateAxisRecord( AxisRecord.AXIS_TYPE_CATEGORY_OR_X_AXIS ) );
		records.Add( CreateBeginRecord() );
		records.Add( CreateCategorySeriesAxisRecord() );
		records.Add( CreateAxisOptionsRecord() );
		records.Add( CreateTickRecord1() );
		records.Add( CreateEndRecord() );
		records.Add( CreateAxisRecord( AxisRecord.AXIS_TYPE_VALUE_AXIS ) );
		records.Add( CreateBeginRecord() );
		records.Add( CreateValueRangeRecord() );
		records.Add( CreateTickRecord2() );
		records.Add( CreateAxisLineFormatRecord( AxisLineFormatRecord.AXIS_TYPE_MAJOR_GRID_LINE ) );
		records.Add( CreateLineFormatRecord(false) );
		records.Add( CreateEndRecord() );
		records.Add( CreatePlotAreaRecord() );
		records.Add( CreateFrameRecord2() );
		records.Add( CreateBeginRecord() );
		records.Add( CreateLineFormatRecord2() );
		records.Add( CreateAreaFormatRecord2() );
		records.Add( CreateEndRecord() );
		records.Add( CreateChartFormatRecord() );
		records.Add( CreateBeginRecord() );
		records.Add( CreateBarRecord() );
		// unknown 1022
		records.Add( CreateLegendRecord() );
		records.Add( CreateBeginRecord() );
		// unknown 104f
		records.Add( CreateTextRecord() );
		records.Add( CreateBeginRecord() );
		// unknown 104f
		records.Add( CreateLinkedDataRecord() );
		records.Add( CreateEndRecord() );
		records.Add( CreateEndRecord() );
		records.Add( CreateEndRecord() );
		records.Add( CreateEndRecord() );
	}

	private LinkedDataRecord CreateLinkedDataRecord()
	{
		LinkedDataRecord r = new LinkedDataRecord();
		r.LinkType=(LinkedDataRecord.LINK_TYPE_TITLE_OR_TEXT);
		r.ReferenceType=(LinkedDataRecord.REFERENCE_TYPE_DIRECT);
		r.IsCustomNumberFormat=(false);
		r.IndexNumberFmtRecord=((short)0);
		r.FormulaOfLink=(null);
		return r;
	}

	private TextRecord CreateTextRecord()
	{
		TextRecord r = new TextRecord();
		r.SetHorizontalAlignment(TextRecord.HORIZONTAL_ALIGNMENT_CENTER);
		r.SetVerticalAlignment(TextRecord.VERTICAL_ALIGNMENT_CENTER);
		r.SetDisplayMode((short)1);
		r.SetRgbColor(0x00000000);
		r.SetX(-37);
		r.SetY(-60);
		r.SetWidth(0);
		r.SetHeight(0);
		r.SetAutoColor(true);
		r.SetShowKey(false);
		r.SetShowValue(false);
		r.SetVertical(false);
		r.SetAutoGeneratedText(true);
		r.SetGenerated(true);
		r.SetAutoLabelDeleted(false);
		r.SetAutoBackground(true);
		r.SetRotation((short)0);
		r.SetShowCategoryLabelAsPercentage(false);
		r.SetShowValueAsPercentage(false);
		r.SetShowBubbleSizes(false);
		r.SetShowLabel(false);
		r.SetIndexOfColorValue((short)77);
		r.SetDataLabelPlacement((short)0);
		r.SetTextRotation((short)0);
		return r;
	}

	private LegendRecord CreateLegendRecord()
	{
		LegendRecord r = new LegendRecord();
		r.SetXAxisUpperLeft(3542);
		r.SetYAxisUpperLeft(1566);
		r.SetXSize(437);
		r.SetYSize(213);
		r.SetType(LegendRecord.TYPE_RIGHT);
		r.SetSpacing(LegendRecord.SPACING_MEDIUM);
		r.SetAutoPosition(true);
		r.SetAutoSeries(true);
		r.SetAutoXPositioning(true);
		r.SetAutoYPositioning(true);
		r.SetVertical(true);
		r.SetDataTable(false);
		return r;
	}

	private BarRecord CreateBarRecord()
	{
		BarRecord r = new BarRecord();
		r.SetBarSpace((short)0);
		r.SetCategorySpace((short)150);
		r.SetHorizontal(false);
		r.SetStacked(false);
		r.SetDisplayAsPercentage(false);
		r.SetShadow(false);
		return r;
	}

	private ChartFormatRecord CreateChartFormatRecord()
	{
		ChartFormatRecord r = new ChartFormatRecord();
		r.SetXPosition(0);
		r.SetYPosition(0);
		r.SetWidth(0);
		r.SetHeight(0);
		r.SetVaryDisplayPattern(false);
		return r;
	}

	private PlotAreaRecord CreatePlotAreaRecord()
	{
		PlotAreaRecord r = new PlotAreaRecord(  );
		return r;
	}

	private AxisLineFormatRecord CreateAxisLineFormatRecord( short format )
	{
		AxisLineFormatRecord r = new AxisLineFormatRecord();
		r.SetAxisType( format );
		return r;
	}

	private ValueRangeRecord CreateValueRangeRecord()
	{
		ValueRangeRecord r = new ValueRangeRecord();
		r.SetMinimumAxisValue( 0.0 );
		r.SetMaximumAxisValue( 0.0 );
		r.SetMajorIncrement( 0 );
		r.SetMinorIncrement( 0 );
		r.SetCategoryAxisCross( 0 );
		r.SetAutomaticMinimum( true );
		r.SetAutomaticMaximum( true );
		r.SetAutomaticMajor( true );
		r.SetAutomaticMinor( true );
		r.SetAutomaticCategoryCrossing( true );
		r.SetLogarithmicScale( false );
		r.SetValuesInReverse( false );
		r.SetCrossCategoryAxisAtMaximum( false );
		r.SetReserved( true );	// what's this do??
		return r;
	}

	private TickRecord CreateTickRecord1()
	{
		TickRecord r = new TickRecord();
		r.SetMajorTickType( (byte) 2 );
		r.SetMinorTickType( (byte) 0 );
		r.SetLabelPosition( (byte) 3 );
		r.SetBackground( (byte) 1 );
		r.SetLabelColorRgb( 0 );
		r.SetZero1( (short) 0 );
		r.SetZero2( (short) 0 );
		r.SetZero3( (short) 45 );
		r.SetAutorotate( true );
		r.SetAutoTextBackground( true );
		r.SetRotation( (short) 0 );
		r.SetAutorotate( true );
		r.SetTickColor( (short) 77 );
		return r;
	}

	private TickRecord CreateTickRecord2()
	{
		TickRecord r = CreateTickRecord1();
		r.SetZero3((short)0);
		return r;
	}

	private AxisOptionsRecord CreateAxisOptionsRecord()
	{
		AxisOptionsRecord r = new AxisOptionsRecord();
		r.SetMinimumCategory( (short) -28644 );
		r.SetMaximumCategory( (short) -28715 );
		r.SetMajorUnitValue( (short) 2 );
		r.SetMajorUnit( (short) 0 );
		r.SetMinorUnitValue( (short) 1 );
		r.SetMinorUnit( (short) 0 );
		r.SetBaseUnit( (short) 0 );
		r.SetCrossingPoint( (short) -28644 );
		r.SetDefaultMinimum( true );
		r.SetDefaultMaximum( true );
		r.SetDefaultMajor( true );
		r.SetDefaultMinorUnit( true );
		r.SetIsDate( true );
		r.SetDefaultBase( true );
		r.SetDefaultCross( true );
		r.SetDefaultDateSettings( true );
		return r;
	}

	private CategorySeriesAxisRecord CreateCategorySeriesAxisRecord()
	{
		CategorySeriesAxisRecord r = new CategorySeriesAxisRecord();
		r.SetCrossingPoint( (short) 1 );
		r.SetLabelFrequency( (short) 1 );
		r.SetTickMarkFrequency( (short) 1 );
		r.SetValueAxisCrossing( true );
		r.SetCrossesFarRight( false );
		r.SetReversed( false );
		return r;
	}

	private AxisRecord CreateAxisRecord( short axisType )
	{
		AxisRecord r = new AxisRecord();
		r.SetAxisType( axisType );
		return r;
	}

	private AxisParentRecord CreateAxisParentRecord()
	{
		AxisParentRecord r = new AxisParentRecord();
		r.SetAxisType( AxisParentRecord.AXIS_TYPE_MAIN );
		r.SetX( 479 );
		r.SetY( 221 );
		r.SetWidth( 2995 );
		r.SetHeight( 2902 );
		return r;
	}

	private AxisUsedRecord CreateAxisUsedRecord( short numAxis )
	{
		AxisUsedRecord r = new AxisUsedRecord();
		r.SetNumAxis( numAxis );
		return r;
	}

	private LinkedDataRecord CreateDirectLinkRecord()
	{
		LinkedDataRecord r = new LinkedDataRecord();
		r.SetLinkType( LinkedDataRecord.LINK_TYPE_TITLE_OR_TEXT );
		r.SetReferenceType( LinkedDataRecord.REFERENCE_TYPE_DIRECT );
		r.SetCustomNumberFormat( false );
		r.SetIndexNumberFmtRecord( (short) 0 );
		r.SetFormulaOfLink(null);
		return r;
	}

	private FontIndexRecord CreateFontIndexRecord( int index )
	{
		FontIndexRecord r = new FontIndexRecord();
		r.SetFontIndex( (short) index );
		return r;
	}

	private TextRecord CreateAllTextRecord()
	{
		TextRecord r = new TextRecord();
		r.SetHorizontalAlignment( TextRecord.HORIZONTAL_ALIGNMENT_CENTER );
		r.SetVerticalAlignment( TextRecord.VERTICAL_ALIGNMENT_CENTER );
		r.SetDisplayMode( TextRecord.DISPLAY_MODE_TRANSPARENT );
		r.SetRgbColor( 0 );
		r.SetX( -37 );
		r.SetY( -60 );
		r.SetWidth( 0 );
		r.SetHeight( 0 );
		r.SetAutoColor( true );
		r.SetShowKey( false );
		r.SetShowValue( true );
		r.SetVertical( false );
		r.SetAutoGeneratedText( true );
		r.SetGenerated( true );
		r.SetAutoLabelDeleted( false );
		r.SetAutoBackground( true );
		r.SetRotation( (short) 0 );
		r.SetShowCategoryLabelAsPercentage( false );
		r.SetShowValueAsPercentage( false );
		r.SetShowBubbleSizes( false );
		r.SetShowLabel( false );
		r.SetIndexOfColorValue( (short) 77 );
		r.SetDataLabelPlacement( (short) 0 );
		r.SetTextRotation( (short) 0 );
		return r;
	}

	private TextRecord CreateUnknownTextRecord()
	{
		TextRecord r = new TextRecord();
		r.SetHorizontalAlignment( TextRecord.HORIZONTAL_ALIGNMENT_CENTER );
		r.SetVerticalAlignment( TextRecord.VERTICAL_ALIGNMENT_CENTER );
		r.SetDisplayMode( TextRecord.DISPLAY_MODE_TRANSPARENT );
		r.SetRgbColor( 0 );
		r.SetX( -37 );
		r.SetY( -60 );
		r.SetWidth( 0 );
		r.SetHeight( 0 );
		r.SetAutoColor( true );
		r.SetShowKey( false );
		r.SetShowValue( false );
		r.SetVertical( false );
		r.SetAutoGeneratedText( true );
		r.SetGenerated( true );
		r.SetAutoLabelDeleted( false );
		r.SetAutoBackground( true );
		r.SetRotation( (short) 0 );
		r.SetShowCategoryLabelAsPercentage( false );
		r.SetShowValueAsPercentage( false );
		r.SetShowBubbleSizes( false );
		r.SetShowLabel( false );
		r.SetIndexOfColorValue( (short) 77 );
		r.SetDataLabelPlacement( (short) 11088 );
		r.SetTextRotation( (short) 0 );
		return r;
	}

	private DefaultDataLabelTextPropertiesRecord CreateDefaultTextRecord( short categoryDataType )
	{
		DefaultDataLabelTextPropertiesRecord r = new DefaultDataLabelTextPropertiesRecord();
		r.SetCategoryDataType( categoryDataType );
		return r;
	}

	private SheetPropertiesRecord CreateSheetPropsRecord()
	{
		SheetPropertiesRecord r = new SheetPropertiesRecord();
		r.SetChartTypeManuallyFormatted( false );
		r.SetPlotVisibleOnly( true );
		r.SetDoNotSizeWithWindow( false );
		r.SetDefaultPlotDimensions( true );
		r.SetAutoPlotArea( false );
		return r;
	}

	private SeriesToChartGroupRecord CreateSeriesToChartGroupRecord()
	{
		return new SeriesToChartGroupRecord();
	}

	private DataFormatRecord CreateDataFormatRecord()
	{
		DataFormatRecord r = new DataFormatRecord();
		r.SetPointNumber( (short) -1 );
		r.SetSeriesIndex( (short) 0 );
		r.SetSeriesNumber( (short) 0 );
		r.SetUseExcel4Colors( false );
		return r;
	}

	private LinkedDataRecord CreateCategoriesLinkedDataRecord()
	{
		LinkedDataRecord r = new LinkedDataRecord();
		r.SetLinkType( LinkedDataRecord.LINK_TYPE_CATEGORIES );
		r.SetReferenceType( LinkedDataRecord.REFERENCE_TYPE_WORKSHEET );
		r.SetCustomNumberFormat( false );
		r.SetIndexNumberFmtRecord( (short) 0 );
		Area3DPtg p = new Area3DPtg(0, 31, 1, 1,
		        false, false, false, false, 0);
		r.SetFormulaOfLink(new Ptg[] { p, });
		return r;
	}

	private LinkedDataRecord CreateValuesLinkedDataRecord()
	{
		LinkedDataRecord r = new LinkedDataRecord();
		r.SetLinkType( LinkedDataRecord.LINK_TYPE_VALUES );
		r.SetReferenceType( LinkedDataRecord.REFERENCE_TYPE_WORKSHEET );
		r.SetCustomNumberFormat( false );
		r.SetIndexNumberFmtRecord( (short) 0 );
		Area3DPtg p = new Area3DPtg(0, 31, 0, 0,
				false, false, false, false, 0);
		r.SetFormulaOfLink(new Ptg[] { p, });
		return r;
	}

	private LinkedDataRecord CreateTitleLinkedDataRecord()
	{
		LinkedDataRecord r = new LinkedDataRecord();
		r.SetLinkType( LinkedDataRecord.LINK_TYPE_TITLE_OR_TEXT );
		r.SetReferenceType( LinkedDataRecord.REFERENCE_TYPE_DIRECT );
		r.SetCustomNumberFormat( false );
		r.SetIndexNumberFmtRecord( (short) 0 );
		r.SetFormulaOfLink(null);
		return r;
	}

	private SeriesRecord CreateSeriesRecord()
	{
		SeriesRecord r = new SeriesRecord();
		r.SetCategoryDataType( SeriesRecord.CATEGORY_DATA_TYPE_NUMERIC );
		r.SetValuesDataType( SeriesRecord.VALUES_DATA_TYPE_NUMERIC );
		r.SetNumCategories( (short) 32 );
		r.SetNumValues( (short) 31 );
		r.SetBubbleSeriesType( SeriesRecord.BUBBLE_SERIES_TYPE_NUMERIC );
		r.SetNumBubbleValues( (short) 0 );
		return r;
	}

	private EndRecord CreateEndRecord()
	{
		return new EndRecord();
	}

	private AreaFormatRecord CreateAreaFormatRecord1()
	{
		AreaFormatRecord r = new AreaFormatRecord();
		r.SetForegroundColor( 16777215 );	 // RGB Color
		r.SetBackgroundColor( 0 );			// RGB Color
		r.SetPattern( (short) 1 );			 // TODO: Add Pattern constants to record
		r.SetAutomatic( true );
		r.SetInvert( false );
		r.SetForecolorIndex( (short) 78 );
		r.SetBackcolorIndex( (short) 77 );
		return r;
	}

	private AreaFormatRecord CreateAreaFormatRecord2()
	{
		AreaFormatRecord r = new AreaFormatRecord();
		r.SetForegroundColor(0x00c0c0c0);
		r.SetBackgroundColor(0x00000000);
		r.SetPattern((short)1);
		r.SetAutomatic(false);
		r.SetInvert(false);
		r.SetForecolorIndex((short)22);
		r.SetBackcolorIndex((short)79);
		return r;
	}

	private LineFormatRecord CreateLineFormatRecord( bool drawTicks )
	{
		LineFormatRecord r = new LineFormatRecord();
		r.SetLineColor( 0 );
		r.SetLinePattern( LineFormatRecord.LINE_PATTERN_SOLID );
		r.SetWeight( (short) -1 );
		r.SetAuto( true );
		r.SetDrawTicks( drawTicks );
		r.SetColourPaletteIndex( (short) 77 );  // what colour is this?
		return r;
	}

	private LineFormatRecord CreateLineFormatRecord2()
	{
		LineFormatRecord r = new LineFormatRecord();
		r.SetLineColor( 0x00808080 );
		r.SetLinePattern( (short) 0 );
		r.SetWeight( (short) 0 );
		r.SetAuto( false );
		r.SetDrawTicks( false );
		r.SetUnknown( false );
		r.SetColourPaletteIndex( (short) 23 );
		return r;
	}

	private FrameRecord CreateFrameRecord1()
	{
		FrameRecord r = new FrameRecord();
		r.SetBorderType( FrameRecord.BORDER_TYPE_REGULAR );
		r.SetAutoSize( false );
		r.SetAutoPosition( true );
		return r;
	}

	private FrameRecord CreateFrameRecord2()
	{
		FrameRecord r = new FrameRecord();
		r.SetBorderType( FrameRecord.BORDER_TYPE_REGULAR );
		r.SetAutoSize( true );
		r.SetAutoPosition( true );
		return r;
	}

	private PlotGrowthRecord CreatePlotGrowthRecord( int horizScale, int vertScale )
	{
		PlotGrowthRecord r = new PlotGrowthRecord();
		r.SetHorizontalScale( horizScale );
		r.SetVerticalScale( vertScale );
		return r;
	}

	private SCLRecord CreateSCLRecord( short numerator, short denominator )
	{
		SCLRecord r = new SCLRecord();
		r.SetDenominator( denominator );
		r.SetNumerator( numerator );
		return r;
	}

	private BeginRecord CreateBeginRecord()
	{
		return new BeginRecord();
	}

	private ChartRecord CreateChartRecord( int x, int y, int width, int height )
	{
		ChartRecord r = new ChartRecord();
		r.SetX( x );
		r.SetY( y );
		r.SetWidth( width );
		r.SetHeight( height );
		return r;
	}

	private UnitsRecord CreateUnitsRecord()
	{
		UnitsRecord r = new UnitsRecord();
		r.SetUnits( (short) 0 );
		return r;
	}


	/**
	 * A series in a chart
	 */
	public class HSSFSeries {
		private SeriesRecord series;
		private SeriesTextRecord seriesTitleText;
		private LinkedDataRecord dataName;
		private LinkedDataRecord dataValues;
		private LinkedDataRecord dataCategoryLabels;
		private LinkedDataRecord dataSecondaryCategoryLabels;

		/* package */ HSSFSeries(SeriesRecord series) {
			this.series = series;
		}

		/* package */ void InsertData(LinkedDataRecord data){
			switch(data.GetLinkType()){
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
		
		/* package */ void SetSeriesTitleText(SeriesTextRecord seriesTitleText)
		{
			this.seriesTitleText = seriesTitleText;
		}
		
		public short GetNumValues() {
			return series.GetNumValues();
		}
		/**
		 * See {@link SeriesRecord}
		 */
		public short GetValueType() {
			return series.GetValuesDataType();
		}

		/**
		 * Returns the series' title, if there is one,
		 *  or null if not
		 */
		public String GetSeriesTitle() {
			if(seriesTitleText != null) {
				return seriesTitleText.GetText();
			}
			return null;
		}

		/**
		 * Changes the series' title, but only if there
		 *  was one already.
		 * TODO - add in the records if not
		 */
		public void SetSeriesTitle(String title) {
			if(seriesTitleText != null) {
				seriesTitleText.SetText(title);
			} else {
				throw new IllegalStateException("No series title found to change");
			}
		}

		/**
		 * @return record with data names
		 */
		public LinkedDataRecord GetDataName(){
			return dataName;
		}
		
		/**
		 * @return record with data values
		 */
		public LinkedDataRecord GetDataValues(){
			return dataValues;
		}
		
		/**
		 * @return record with data category labels
		 */
		public LinkedDataRecord GetDataCategoryLabels(){
			return dataCategoryLabels;
		}
		
		/**
		 * @return record with data secondary category labels
		 */
		public LinkedDataRecord GetDataSecondaryCategoryLabels() {
			return dataSecondaryCategoryLabels;
		}
		
		/**
		 * @return record with series
		 */
		public SeriesRecord GetSeries() {
			return series;
		}
		
		private CellRangeAddressBase GetCellRange(LinkedDataRecord linkedDataRecord) {
			if (linkedDataRecord == null)
			{
				return null ;
			}
			
			int firstRow = 0;
			int lastRow = 0;
			int firstCol = 0;
			int lastCol = 0;
			
			for (Ptg ptg : linkedDataRecord.GetFormulaOfLink()) {
				if (ptg is AreaPtgBase) {
					AreaPtgBase areaPtg = (AreaPtgBase) ptg;
					
					firstRow = areaPtg.GetFirstRow();
					lastRow = areaPtg.GetLastRow();
					
					firstCol = areaPtg.GetFirstColumn();
					lastCol = areaPtg.GetLastColumn();
				}
			}
			
			return new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		}
		
		public CellRangeAddressBase GetValuesCellRange() {
			return GetCellRange(dataValues);
		}
	
		public CellRangeAddressBase GetCategoryLabelsCellRange() {
			return GetCellRange(dataCategoryLabels);
		}
	
		private Integer SetVerticalCellRange(LinkedDataRecord linkedDataRecord,
				                             CellRangeAddressBase range) {
			if (linkedDataRecord == null)
			{
				return null;
			}
			
			List<Ptg> ptgList = new ArrayList<Ptg>();
			
			int rowCount = (range.GetLastRow() - range.GetFirstRow()) + 1;
			int colCount = (range.GetLastColumn() - range.GetFirstColumn()) + 1;
			
			for (Ptg ptg : linkedDataRecord.GetFormulaOfLink()) {
				if (ptg is AreaPtgBase) {
					AreaPtgBase areaPtg = (AreaPtgBase) ptg;
					
					areaPtg.SetFirstRow(range.GetFirstRow());
					areaPtg.SetLastRow(range.GetLastRow());
					
					areaPtg.SetFirstColumn(range.GetFirstColumn());
					areaPtg.SetLastColumn(range.GetLastColumn());
					ptgList.Add(areaPtg);
				}
			}
			
			linkedDataRecord.SetFormulaOfLink(ptgList.ToArray(new Ptg[ptgList.Size()]));
			
			return rowCount * colCount;
		}
		
		public void SetValuesCellRange(CellRangeAddressBase range) {
			Integer count = SetVerticalCellRange(dataValues, range);
			if (count == null)
			{
				return;
			}
			
			series.SetNumValues((short)(int)count);
		}
		
		public void SetCategoryLabelsCellRange(CellRangeAddressBase range) {
			Integer count = SetVerticalCellRange(dataCategoryLabels, range);
			if (count == null)
			{
				return;
			}
			
			series.SetNumCategories((short)(int)count);
		}
	}
	
	public HSSFSeries CreateSeries() {
        ArrayList<RecordBase> seriesTemplate = new ArrayList<RecordBase>();
		bool seriesTemplateFilled = false;
		
		int idx = 0;
		int deep = 0;
		int chartRecordIdx = -1;
		int chartDeep = -1;
		int lastSeriesDeep = -1;
		int endSeriesRecordIdx = -1;
		int seriesIdx = 0;
		 List<RecordBase> records = sheet.GetSheet().GetRecords();
		
		/* store first series as template and find last series index */
		for( RecordBase record : records) {		
			
			idx++;
			
			if (record is BeginRecord) {
				deep++;
			} else if (record is EndRecord) {
				deep--;
				
				if (lastSeriesDeep == deep) {
					lastSeriesDeep = -1;
					endSeriesRecordIdx = idx;
					if (!seriesTemplateFilled) {
						seriesTemplate.Add(record);
						seriesTemplateFilled = true;
					}
				}
				
				if (chartDeep == deep) {
					break;
				}
			}
			
			if (record is ChartRecord) {
				if (record == chartRecord) {
					chartRecordIdx = idx;
					chartDeep = deep;
				}
			} else if (record is SeriesRecord) {
				if (chartRecordIdx != -1) {
					seriesIdx++;
					lastSeriesDeep = deep;
				}
			}
			
			if (lastSeriesDeep != -1 && !seriesTemplateFilled) {
				seriesTemplate.Add(record) ;
			}
		}
		
		/* check if a series was found */
		if (endSeriesRecordIdx == -1) {
			return null;
		}
		
		/* next index in the records list where the new series can be inserted */
		idx = endSeriesRecordIdx + 1;

		HSSFSeries newSeries = null;
		
		/* duplicate record of the template series */
		ArrayList<RecordBase> clonedRecords = new ArrayList<RecordBase>();
		for( RecordBase record : seriesTemplate) {		
			
			Record newRecord = null;
			
			if (record is BeginRecord) {
				newRecord = new BeginRecord();
			} else if (record is EndRecord) {
				newRecord = new EndRecord();
			} else if (record is SeriesRecord) {
				SeriesRecord seriesRecord = (SeriesRecord) ((SeriesRecord)record).Clone();
				newSeries = new HSSFSeries(seriesRecord);
				newRecord = seriesRecord;
			} else if (record is LinkedDataRecord) {
				LinkedDataRecord linkedDataRecord = (LinkedDataRecord) ((LinkedDataRecord)record).Clone();
				if (newSeries != null) {
					newSeries.InsertData(linkedDataRecord);
				}
				newRecord = linkedDataRecord;
			} else if (record is DataFormatRecord) {
				DataFormatRecord dataFormatRecord = (DataFormatRecord) ((DataFormatRecord)record).Clone();
				
				dataFormatRecord.SetSeriesIndex((short)seriesIdx) ;
				dataFormatRecord.SetSeriesNumber((short)seriesIdx) ;
				
				newRecord = dataFormatRecord;
			} else if (record is SeriesTextRecord) {
				SeriesTextRecord seriesTextRecord = (SeriesTextRecord) ((SeriesTextRecord)record).Clone();
				if (newSeries != null) {
					newSeries.SetSeriesTitleText(seriesTextRecord);
				}
				newRecord = seriesTextRecord;
			} else if (record is Record) {
				newRecord = (Record) ((Record)record).Clone();
			}
			
			if (newRecord != null)
			{
				clonedRecords.Add(newRecord);
			}
		}
		
		/* check if a user model series object was created */
		if (newSeries == null)
		{
			return null;
		}
		
		/* transfer series to record list */
		for( RecordBase record : clonedRecords) {		
			records.Add(idx++, record);
		}
		
		return newSeries;
	}
	
	public bool RemoveSeries(HSSFSeries series) {
		int idx = 0;
		int deep = 0;
		int chartDeep = -1;
		int lastSeriesDeep = -1;
		int seriesIdx = -1;
		bool removeSeries = false;
		bool chartEntered = false;
		bool result = false;
		 List<RecordBase> records = sheet.GetSheet().GetRecords();
		
		/* store first series as template and find last series index */
		Iterator<RecordBase> iter = records.Iterator();
		while (iter.HasNext()) {		
			RecordBase record = iter.Next();
			idx++;
			
			if (record is BeginRecord) {
				deep++;
			} else if (record is EndRecord) {
				deep--;
				
				if (lastSeriesDeep == deep) {
					lastSeriesDeep = -1;
					
					if (removeSeries) {
						removeSeries = false;
						result = true;
						iter.Remove();
					}
				}
				
				if (chartDeep == deep) {
					break;
				}
			}
			
			if (record is ChartRecord) {
				if (record == chartRecord) {
					chartDeep = deep;
					chartEntered = true;
				}
			} else if (record is SeriesRecord) {
				if (chartEntered) {
					if (series.series == record) {
						lastSeriesDeep = deep;
						removeSeries = true;
					} else {
						seriesIdx++;
					}
				}
			} else if (record is DataFormatRecord) {
				if (chartEntered && !removeSeries) {
					DataFormatRecord dataFormatRecord = (DataFormatRecord) record;
					dataFormatRecord.SetSeriesIndex((short) seriesIdx);
					dataFormatRecord.SetSeriesNumber((short) seriesIdx);
				}
			}
			
			if (removeSeries) {
				iter.Remove();
			}
		}
		
		return result;
	}
	
	public HSSFChartType GetType() {
		return type;
	}
}
}

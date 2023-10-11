
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
/* ================================================================
 * About NPOI
 * POI Version: 3.8 beta4
 * Date: 2012-02-15
 * 
 * ==============================================================*/

namespace NPOI.HSSF.Record
{
    using System;
    using System.IO;
        using System.Collections;
    using System.Collections.Generic;
    using NPOI.HSSF.Record.Chart;
    using NPOI.HSSF.Record.PivotTable;
    using NPOI.HSSF.Record.AutoFilter;
    using NPOI.Util;
    using System.Globalization;
    using System.Linq;

    /**
     * Title:  Record Factory
     * Description:  Takes a stream and outputs an array of Record objects.
     *
     * @deprecated use {@link org.apache.poi.hssf.eventmodel.EventRecordFactory} instead
     * @see org.apache.poi.hssf.eventmodel.EventRecordFactory
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Csaba Nagy (ncsaba at yahoo dot com)
     */

    // REMOVE-REFLECTION: Reflection is removed by turning into a lookup.
    public class RecordFactory
    {
        private static int NUM_RECORDS = 512;
        private static Type[] recordClasses;

        private static Dictionary<short, Type> _recordTypes;

        static RecordFactory()
        {
            // For reference only

            /*recordClasses = new Type[]
            {
		        typeof(ArrayRecord),
                typeof(AutoFilterInfoRecord),
		        typeof(BackupRecord),
		        typeof(BlankRecord),
		        typeof(BOFRecord),
		        typeof(BookBoolRecord),
		        typeof(BoolErrRecord),
		        typeof(BottomMarginRecord),
		        typeof(BoundSheetRecord),
		        typeof(CalcCountRecord),
		        typeof(CalcModeRecord),
		        typeof(CFHeaderRecord),
                typeof(CFHeader12Record),
                typeof(CFRuleRecord),
                typeof(CFRule12Record),
                typeof(ChartRecord),
                typeof(AlRunsRecord),
                //typeof(CodeNameRecord),
		        typeof(CodepageRecord),
		        typeof(ColumnInfoRecord),
		        typeof(ContinueRecord),
		        typeof(CountryRecord),
		        typeof(CRNCountRecord),
		        typeof(CRNRecord),
		        typeof(DateWindow1904Record),
		        typeof(DBCellRecord),
                typeof(DConRefRecord),
		        typeof(DefaultColWidthRecord),
		        typeof(DefaultRowHeightRecord),
		        typeof(DeltaRecord),
		        typeof(DimensionsRecord),
		        typeof(DrawingGroupRecord),
		        typeof(DrawingRecord),
		        typeof(DrawingSelectionRecord),
		        typeof(DSFRecord),
		        typeof(DVALRecord),
		        typeof(DVRecord),
		        typeof(EOFRecord),
		        typeof(ExtendedFormatRecord),
		        typeof(ExternalNameRecord),
		        typeof(ExternSheetRecord),
		        typeof(ExtSSTRecord),
		        typeof(FilePassRecord),
		        typeof(FileSharingRecord),
		        typeof(FnGroupCountRecord),
		        typeof(FontRecord),
		        typeof(FooterRecord),
		        typeof(FormatRecord),
		        typeof(FormulaRecord),
		        typeof(GridsetRecord),
		        typeof(GutsRecord),
		        typeof(HCenterRecord),
		        typeof(HeaderRecord),
                typeof(HeaderFooterRecord),
		        typeof(HideObjRecord),
		        typeof(HorizontalPageBreakRecord),
		        typeof(HyperlinkRecord),
		        typeof(IndexRecord),
		        typeof(InterfaceEndRecord),
		        typeof(InterfaceHdrRecord),
		        typeof(IterationRecord),
		        typeof(LabelRecord),
		        typeof(LabelSSTRecord),
		        typeof(LeftMarginRecord),
		        typeof(MergeCellsRecord),
		        typeof(MMSRecord),
		        typeof(MulBlankRecord),
		        typeof(MulRKRecord),
		        typeof(NameRecord),
                typeof(NameCommentRecord),
		        typeof(NoteRecord),
		        typeof(NumberRecord),
		        typeof(ObjectProtectRecord),
		        typeof(ObjRecord),
		        typeof(PaletteRecord),
		        typeof(PaneRecord),
		        typeof(PasswordRecord),
		        typeof(PasswordRev4Record),
		        typeof(PrecisionRecord),
		        typeof(PrintGridlinesRecord),
		        typeof(PrintHeadersRecord),
		        typeof(PrintSetupRecord),
                typeof(PrintSizeRecord),
		        typeof(ProtectionRev4Record),
		        typeof(ProtectRecord),
		        typeof(RecalcIdRecord),
		        typeof(RefModeRecord),
		        typeof(RefreshAllRecord),
		        typeof(RightMarginRecord),
		        typeof(RKRecord),
		        typeof(RowRecord),
		        typeof(SaveRecalcRecord),
		        typeof(ScenarioProtectRecord),
                typeof(SCLRecord),
		        typeof(SelectionRecord),
                typeof(SeriesRecord),
                typeof(SeriesTextRecord),
		        typeof(SharedFormulaRecord),
		        typeof(SSTRecord),
		        typeof(StringRecord),
		        typeof(StyleRecord),
		        typeof(SupBookRecord),
		        typeof(TabIdRecord),
		        typeof(TableRecord),
                typeof(TableStylesRecord),
		        typeof(TextObjectRecord),
		        typeof(TopMarginRecord),
		        typeof(UncalcedRecord),
		        typeof(UseSelFSRecord),
                typeof(UserSViewBegin),
		        typeof(UserSViewEnd),
                typeof(ValueRangeRecord),
		        typeof(VCenterRecord),
		        typeof(VerticalPageBreakRecord),
		        typeof(WindowOneRecord),
		        typeof(WindowProtectRecord),
		        typeof(WindowTwoRecord),
		        typeof(WriteAccessRecord),
		        typeof(WriteProtectRecord),
		        typeof(WSBoolRecord),
                typeof(SheetExtRecord),

                // chart records
                //typeof(AlRunsRecord),
                typeof(AreaFormatRecord),
                typeof(AreaRecord),
                typeof(AttachedLabelRecord),
                typeof(AxcExtRecord),
                typeof(AxisLineFormatRecord),
                typeof(AxisParentRecord),
                typeof(AxisRecord),
                typeof(AxesUsedRecord),
                typeof(BarRecord),
                typeof(BeginRecord),
                typeof(BopPopCustomRecord),
                typeof(BopPopRecord),
                typeof(CatLabRecord),
                typeof(CatSerRangeRecord),
                typeof(Chart3DBarShapeRecord),
                typeof(Chart3dRecord),
                typeof(ChartEndObjectRecord),
                typeof(ChartFormatRecord),
                typeof(ChartFRTInfoRecord),
                //typeof(ChartRecord),
                typeof(ChartStartObjectRecord),
                typeof(CrtLayout12ARecord),
                typeof(CrtLayout12Record),
                typeof(CrtLineRecord),
                typeof(CrtLinkRecord),
                typeof(CrtMlFrtContinueRecord),
                typeof(CrtMlFrtRecord),
                typeof(DataFormatRecord),
                typeof(DataLabExtContentsRecord),
                typeof(DataLabExtRecord),
                typeof(DatRecord),
                typeof(DefaultTextRecord),
                typeof(DropBarRecord),
                typeof(EndBlockRecord),
                typeof(EndRecord),
                typeof(LinkedDataRecord),
                typeof(Fbi2Record),
                typeof(FbiRecord),
                typeof(FontIndexRecord),
                typeof(FrameRecord),
                typeof(FrtFontListRecord),
                typeof(GelFrameRecord),
                typeof(IFmtRecordRecord),
                typeof(LegendExceptionRecord),
                typeof(LegendRecord),
                typeof(LineFormatRecord),
                typeof(MarkerFormatRecord),
                typeof(ObjectLinkRecord),
                typeof(PicFRecord),
                typeof(PieFormatRecord),
                typeof(PieRecord),
                typeof(PlotAreaRecord),
                typeof(PlotGrowthRecord),
                typeof(PosRecord),
                typeof(RadarAreaRecord),
                typeof(RadarRecord),
                typeof(RichTextStreamRecord),
                typeof(ScatterRecord),
                typeof(SerAuxErrBarRecord),
                typeof(SerAuxTrendRecord),
                typeof(SerFmtRecord),
                typeof(SeriesIndexRecord),
                typeof(SeriesListRecord),
                //typeof(SeriesRecord),
                //typeof(SeriesTextRecord),
                //typeof(SeriesToChartGroupRecord),
                typeof(SerParentRecord),
                typeof(SerToCrtRecord),
                typeof(ShapePropsStreamRecord),
                typeof(ShtPropsRecord),
                typeof(StartBlockRecord),
                typeof(SurfRecord),
                typeof(TextPropsStreamRecord),
                typeof(TextRecord),
                typeof(TickRecord),
                typeof(UnitsRecord),
                //typeof(ValueRangeRecord),
                typeof(YMultRecord),
                		        


                //typeof(BeginRecord),
                //typeof(ChartFRTInfoRecord),
                //typeof(StartBlockRecord),
                //typeof(EndBlockRecord),
                //typeof(ChartStartObjectRecord),
                //typeof(ChartEndObjectRecord),
                //typeof(CatLabRecord),
                //typeof(EndRecord),
                //typeof(PrintSizeRecord),
                
                //typeof(AreaFormatRecord),
                //typeof(AreaRecord),
                //typeof(AxisLineRecord),
                //typeof(AxcExtRecord),
                //typeof(AxisParentRecord),
                //typeof(AxisRecord),
                //typeof(Chart3DBarShapeRecord),
                //typeof(CrtLinkRecord),
                //typeof(AxisUsedRecord),
                //typeof(BarRecord),
                //typeof(CatSerRangeRecord),
                //typeof(DataFormatRecord),
                //typeof(DataLabExtRecord),
                //typeof(DefaultTextRecord),
                //typeof(FrameRecord),
                //typeof(LegendRecord),
                //typeof(FbiRecord),
                //typeof(FontXRecord),
                //typeof(LineFormatRecord),
                //typeof(BRAIRecord),
                //typeof(IFmtRecordRecord),
                //typeof(ObjectLinkRecord),
                //typeof(PlotAreaRecord),
                //typeof(PlotGrowthRecord),
                //typeof(PosRecord),
                //typeof(SCLRecord),
                //typeof(SerToCrtRecord),
                ////typeof(SeriesIndexRecord),
                //typeof(AttachedLabelRecord),
                //typeof(SeriesListRecord),
                ////typeof(SeriesToChartGroupRecord),
                //typeof(ShtPropsRecord),
                //typeof(TextRecord),
                //typeof(TickRecord),
                //typeof(UnitsRecord),

                //typeof(CrtLayout12Record),
                //typeof(Chart3dRecord),
                //typeof(PieRecord),
                //typeof(PieFormatRecord),
                //typeof(CrtLayout12ARecord),
                //typeof(MarkerFormatRecord),
                //typeof(ChartFormatRecord),
                //typeof(SeriesIndexRecord),
                //typeof(GelFrameRecord),
                //typeof(PicFRecord),
                //typeof(CrtMlFrtRecord),
                //typeof(CrtMlFrtContinueRecord),
                //typeof(ShapePropsStreamRecord),
                //end

          		// pivot table records
		        typeof(DataItemRecord),
		        typeof(ExtendedPivotTableViewFieldsRecord),
		        typeof(PageItemRecord),
		        typeof(StreamIDRecord),
		        typeof(ViewDefinitionRecord), 
		        typeof(ViewFieldsRecord),
		        typeof(ViewSourceRecord),

                //autofilter
                typeof(AutoFilterRecord),
                typeof(FilterModeRecord),

                typeof(Excel9FileRecord)
            };*/

            _recordTypes = new Dictionary<short, Type>
            {
                { ArrayRecord.sid, typeof(ArrayRecord) },
                { AutoFilterInfoRecord.sid, typeof(AutoFilterInfoRecord) },
                { BackupRecord.sid, typeof(BackupRecord) },
                { BlankRecord.sid, typeof(BlankRecord) },
                { BOFRecord.sid, typeof(BOFRecord) },
                { BookBoolRecord.sid, typeof(BookBoolRecord) },
                { BoolErrRecord.sid, typeof(BoolErrRecord) },
                { BottomMarginRecord.sid, typeof(BottomMarginRecord) },
                { BoundSheetRecord.sid, typeof(BoundSheetRecord) },
                { CalcCountRecord.sid, typeof(CalcCountRecord) },
                { CalcModeRecord.sid, typeof(CalcModeRecord) },
                { CFHeaderRecord.sid, typeof(CFHeaderRecord) },
                { CFHeader12Record.sid, typeof(CFHeader12Record) },
                { CFRuleRecord.sid, typeof(CFRuleRecord) },
                { CFRule12Record.sid, typeof(CFRule12Record) },
                { ChartRecord.sid, typeof(ChartRecord) },
                { AlRunsRecord.sid, typeof(AlRunsRecord) },
                { CodepageRecord.sid, typeof(CodepageRecord) },
                { ColumnInfoRecord.sid, typeof(ColumnInfoRecord) },
                { ContinueRecord.sid, typeof(ContinueRecord) },
                { CountryRecord.sid, typeof(CountryRecord) },
                { CRNCountRecord.sid, typeof(CRNCountRecord) },
                { CRNRecord.sid, typeof(CRNRecord) },
                { DateWindow1904Record.sid, typeof(DateWindow1904Record) },
                { DBCellRecord.sid, typeof(DBCellRecord) },
                { DConRefRecord.sid, typeof(DConRefRecord) },
                { DefaultColWidthRecord.sid, typeof(DefaultColWidthRecord) },
                { DefaultRowHeightRecord.sid, typeof(DefaultRowHeightRecord) },
                { DeltaRecord.sid, typeof(DeltaRecord) },
                { DimensionsRecord.sid, typeof(DimensionsRecord) },
                { DrawingGroupRecord.sid, typeof(DrawingGroupRecord) },
                { DrawingRecord.sid, typeof(DrawingRecord) },
                { DrawingSelectionRecord.sid, typeof(DrawingSelectionRecord) },
                { DSFRecord.sid, typeof(DSFRecord) },
                { DVALRecord.sid, typeof(DVALRecord) },
                { DVRecord.sid, typeof(DVRecord) },
                { EOFRecord.sid, typeof(EOFRecord) },
                { ExtendedFormatRecord.sid, typeof(ExtendedFormatRecord) },
                { ExternalNameRecord.sid, typeof(ExternalNameRecord) },
                { ExternSheetRecord.sid, typeof(ExternSheetRecord) },
                { ExtSSTRecord.sid, typeof(ExtSSTRecord) },
                { FilePassRecord.sid, typeof(FilePassRecord) },
                { FileSharingRecord.sid, typeof(FileSharingRecord) },
                { FnGroupCountRecord.sid, typeof(FnGroupCountRecord) },
                { FontRecord.sid, typeof(FontRecord) },
                { FooterRecord.sid, typeof(FooterRecord) },
                { FormatRecord.sid, typeof(FormatRecord) },
                { FormulaRecord.sid, typeof(FormulaRecord) },
                { GridsetRecord.sid, typeof(GridsetRecord) },
                { GutsRecord.sid, typeof(GutsRecord) },
                { HCenterRecord.sid, typeof(HCenterRecord) },
                { HeaderRecord.sid, typeof(HeaderRecord) },
                { HeaderFooterRecord.sid, typeof(HeaderFooterRecord) },
                { HideObjRecord.sid, typeof(HideObjRecord) },
                { HorizontalPageBreakRecord.sid, typeof(HorizontalPageBreakRecord) },
                { HyperlinkRecord.sid, typeof(HyperlinkRecord) },
                { IndexRecord.sid, typeof(IndexRecord) },
                { InterfaceEndRecord.sid, typeof(InterfaceEndRecord) },
                { InterfaceHdrRecord.sid, typeof(InterfaceHdrRecord) },
                { IterationRecord.sid, typeof(IterationRecord) },
                { LabelRecord.sid, typeof(LabelRecord) },
                { LabelSSTRecord.sid, typeof(LabelSSTRecord) },
                { LeftMarginRecord.sid, typeof(LeftMarginRecord) },
                { MergeCellsRecord.sid, typeof(MergeCellsRecord) },
                { MMSRecord.sid, typeof(MMSRecord) },
                { MulBlankRecord.sid, typeof(MulBlankRecord) },
                { MulRKRecord.sid, typeof(MulRKRecord) },
                { NameRecord.sid, typeof(NameRecord) },
                { NameCommentRecord.sid, typeof(NameCommentRecord) },
                { NoteRecord.sid, typeof(NoteRecord) },
                { NumberRecord.sid, typeof(NumberRecord) },
                { ObjectProtectRecord.sid, typeof(ObjectProtectRecord) },
                { ObjRecord.sid, typeof(ObjRecord) },
                { PaletteRecord.sid, typeof(PaletteRecord) },
                { PaneRecord.sid, typeof(PaneRecord) },
                { PasswordRecord.sid, typeof(PasswordRecord) },
                { PasswordRev4Record.sid, typeof(PasswordRev4Record) },
                { PrecisionRecord.sid, typeof(PrecisionRecord) },
                { PrintGridlinesRecord.sid, typeof(PrintGridlinesRecord) },
                { PrintHeadersRecord.sid, typeof(PrintHeadersRecord) },
                { PrintSetupRecord.sid, typeof(PrintSetupRecord) },
                { PrintSizeRecord.sid, typeof(PrintSizeRecord) },
                { ProtectionRev4Record.sid, typeof(ProtectionRev4Record) },
                { ProtectRecord.sid, typeof(ProtectRecord) },
                { RecalcIdRecord.sid, typeof(RecalcIdRecord) },
                { RefModeRecord.sid, typeof(RefModeRecord) },
                { RefreshAllRecord.sid, typeof(RefreshAllRecord) },
                { RightMarginRecord.sid, typeof(RightMarginRecord) },
                { RKRecord.sid, typeof(RKRecord) },
                { RowRecord.sid, typeof(RowRecord) },
                { SaveRecalcRecord.sid, typeof(SaveRecalcRecord) },
                { ScenarioProtectRecord.sid, typeof(ScenarioProtectRecord) },
                { SCLRecord.sid, typeof(SCLRecord) },
                { SelectionRecord.sid, typeof(SelectionRecord) },
                { SeriesRecord.sid, typeof(SeriesRecord) },
                { SeriesTextRecord.sid, typeof(SeriesTextRecord) },
                { SharedFormulaRecord.sid, typeof(SharedFormulaRecord) },
                { SSTRecord.sid, typeof(SSTRecord) },
                { StringRecord.sid, typeof(StringRecord) },
                { StyleRecord.sid, typeof(StyleRecord) },
                { SupBookRecord.sid, typeof(SupBookRecord) },
                { TabIdRecord.sid, typeof(TabIdRecord) },
                { TableRecord.sid, typeof(TableRecord) },
                { TableStylesRecord.sid, typeof(TableStylesRecord) },
                { TextObjectRecord.sid, typeof(TextObjectRecord) },
                { TopMarginRecord.sid, typeof(TopMarginRecord) },
                { UncalcedRecord.sid, typeof(UncalcedRecord) },
                { UseSelFSRecord.sid, typeof(UseSelFSRecord) },
                { UserSViewBegin.sid, typeof(UserSViewBegin) },
                { UserSViewEnd.sid, typeof(UserSViewEnd) },
                { ValueRangeRecord.sid, typeof(ValueRangeRecord) },
                { VCenterRecord.sid, typeof(VCenterRecord) },
                { VerticalPageBreakRecord.sid, typeof(VerticalPageBreakRecord) },
                { WindowOneRecord.sid, typeof(WindowOneRecord) },
                { WindowProtectRecord.sid, typeof(WindowProtectRecord) },
                { WindowTwoRecord.sid, typeof(WindowTwoRecord) },
                { WriteAccessRecord.sid, typeof(WriteAccessRecord) },
                { WriteProtectRecord.sid, typeof(WriteProtectRecord) },
                { WSBoolRecord.sid, typeof(WSBoolRecord) },
                { SheetExtRecord.sid, typeof(SheetExtRecord) },
                { AreaFormatRecord.sid, typeof(AreaFormatRecord) },
                { AreaRecord.sid, typeof(AreaRecord) },
                { AttachedLabelRecord.sid, typeof(AttachedLabelRecord) },
                { AxcExtRecord.sid, typeof(AxcExtRecord) },
                { AxisLineFormatRecord.sid, typeof(AxisLineFormatRecord) },
                { AxisParentRecord.sid, typeof(AxisParentRecord) },
                { AxisRecord.sid, typeof(AxisRecord) },
                { AxesUsedRecord.sid, typeof(AxesUsedRecord) },
                { BarRecord.sid, typeof(BarRecord) },
                { BeginRecord.sid, typeof(BeginRecord) },
                { BopPopCustomRecord.sid, typeof(BopPopCustomRecord) },
                { BopPopRecord.sid, typeof(BopPopRecord) },
                { CatLabRecord.sid, typeof(CatLabRecord) },
                { CatSerRangeRecord.sid, typeof(CatSerRangeRecord) },
                { Chart3DBarShapeRecord.sid, typeof(Chart3DBarShapeRecord) },
                { Chart3dRecord.sid, typeof(Chart3dRecord) },
                { ChartEndObjectRecord.sid, typeof(ChartEndObjectRecord) },
                { ChartFormatRecord.sid, typeof(ChartFormatRecord) },
                { ChartFRTInfoRecord.sid, typeof(ChartFRTInfoRecord) },
                { ChartStartObjectRecord.sid, typeof(ChartStartObjectRecord) },
                { CrtLayout12ARecord.sid, typeof(CrtLayout12ARecord) },
                { CrtLayout12Record.sid, typeof(CrtLayout12Record) },
                { CrtLineRecord.sid, typeof(CrtLineRecord) },
                { CrtLinkRecord.sid, typeof(CrtLinkRecord) },
                { CrtMlFrtContinueRecord.sid, typeof(CrtMlFrtContinueRecord) },
                { CrtMlFrtRecord.sid, typeof(CrtMlFrtRecord) },
                { DataFormatRecord.sid, typeof(DataFormatRecord) },
                { DataLabExtContentsRecord.sid, typeof(DataLabExtContentsRecord) },
                { DataLabExtRecord.sid, typeof(DataLabExtRecord) },
                { DatRecord.sid, typeof(DatRecord) },
                { DefaultTextRecord.sid, typeof(DefaultTextRecord) },
                { DropBarRecord.sid, typeof(DropBarRecord) },
                { EndBlockRecord.sid, typeof(EndBlockRecord) },
                { EndRecord.sid, typeof(EndRecord) },
                { LinkedDataRecord.sid, typeof(LinkedDataRecord) },
                { Fbi2Record.sid, typeof(Fbi2Record) },
                { FbiRecord.sid, typeof(FbiRecord) },
                { FontIndexRecord.sid, typeof(FontIndexRecord) },
                { FrameRecord.sid, typeof(FrameRecord) },
                { FrtFontListRecord.sid, typeof(FrtFontListRecord) },
                { GelFrameRecord.sid, typeof(GelFrameRecord) },
                { IFmtRecordRecord.sid, typeof(IFmtRecordRecord) },
                { LegendExceptionRecord.sid, typeof(LegendExceptionRecord) },
                { LegendRecord.sid, typeof(LegendRecord) },
                { LineFormatRecord.sid, typeof(LineFormatRecord) },
                { MarkerFormatRecord.sid, typeof(MarkerFormatRecord) },
                { ObjectLinkRecord.sid, typeof(ObjectLinkRecord) },
                { PicFRecord.sid, typeof(PicFRecord) },
                { PieFormatRecord.sid, typeof(PieFormatRecord) },
                { PieRecord.sid, typeof(PieRecord) },
                { PlotAreaRecord.sid, typeof(PlotAreaRecord) },
                { PlotGrowthRecord.sid, typeof(PlotGrowthRecord) },
                { PosRecord.sid, typeof(PosRecord) },
                { RadarAreaRecord.sid, typeof(RadarAreaRecord) },
                { RadarRecord.sid, typeof(RadarRecord) },
                { RichTextStreamRecord.sid, typeof(RichTextStreamRecord) },
                { ScatterRecord.sid, typeof(ScatterRecord) },
                { SerAuxErrBarRecord.sid, typeof(SerAuxErrBarRecord) },
                { SerAuxTrendRecord.sid, typeof(SerAuxTrendRecord) },
                { SerFmtRecord.sid, typeof(SerFmtRecord) },
                { SeriesIndexRecord.sid, typeof(SeriesIndexRecord) },
                { SeriesListRecord.sid, typeof(SeriesListRecord) },
                { SerParentRecord.sid, typeof(SerParentRecord) },
                { SerToCrtRecord.sid, typeof(SerToCrtRecord) },
                { ShapePropsStreamRecord.sid, typeof(ShapePropsStreamRecord) },
                { ShtPropsRecord.sid, typeof(ShtPropsRecord) },
                { StartBlockRecord.sid, typeof(StartBlockRecord) },
                { SurfRecord.sid, typeof(SurfRecord) },
                { TextPropsStreamRecord.sid, typeof(TextPropsStreamRecord) },
                { TextRecord.sid, typeof(TextRecord) },
                { TickRecord.sid, typeof(TickRecord) },
                { UnitsRecord.sid, typeof(UnitsRecord) },
                { YMultRecord.sid, typeof(YMultRecord) },
                { DataItemRecord.sid, typeof(DataItemRecord) },
                { ExtendedPivotTableViewFieldsRecord.sid, typeof(ExtendedPivotTableViewFieldsRecord) },
                { PageItemRecord.sid, typeof(PageItemRecord) },
                { StreamIDRecord.sid, typeof(StreamIDRecord) },
                { ViewDefinitionRecord.sid, typeof(ViewDefinitionRecord) },
                { ViewFieldsRecord.sid, typeof(ViewFieldsRecord) },
                { ViewSourceRecord.sid, typeof(ViewSourceRecord) },
                { AutoFilterRecord.sid, typeof(AutoFilterRecord) },
                { FilterModeRecord.sid, typeof(FilterModeRecord) },
                { Excel9FileRecord.sid, typeof(Excel9FileRecord) }
            };
            // recordTypes.Add()
        }

        private static Record CreateBySid(short sid, RecordInputStream stream)
        {
            return sid switch
            {
                ArrayRecord.sid => new ArrayRecord(stream),
                AutoFilterInfoRecord.sid => new AutoFilterInfoRecord(stream),
                BackupRecord.sid => new BackupRecord(stream),
                BlankRecord.sid => new BlankRecord(stream),
                BOFRecord.sid => new BOFRecord(stream),
                BookBoolRecord.sid => new BookBoolRecord(stream),
                BoolErrRecord.sid => new BoolErrRecord(stream),
                BottomMarginRecord.sid => new BottomMarginRecord(stream),
                BoundSheetRecord.sid => new BoundSheetRecord(stream),
                CalcCountRecord.sid => new CalcCountRecord(stream),
                CalcModeRecord.sid => new CalcModeRecord(stream),
                CFHeaderRecord.sid => new CFHeaderRecord(stream),
                CFHeader12Record.sid => new CFHeader12Record(stream),
                CFRuleRecord.sid => new CFRuleRecord(stream),
                CFRule12Record.sid => new CFRule12Record(stream),
                ChartRecord.sid => new ChartRecord(stream),
                AlRunsRecord.sid => new AlRunsRecord(stream),
                CodepageRecord.sid => new CodepageRecord(stream),
                ColumnInfoRecord.sid => new ColumnInfoRecord(stream),
                ContinueRecord.sid => new ContinueRecord(stream),
                CountryRecord.sid => new CountryRecord(stream),
                CRNCountRecord.sid => new CRNCountRecord(stream),
                CRNRecord.sid => new CRNRecord(stream),
                DateWindow1904Record.sid => new DateWindow1904Record(stream),
                DBCellRecord.sid => new DBCellRecord(stream),
                DConRefRecord.sid => new DConRefRecord(stream),
                DefaultColWidthRecord.sid => new DefaultColWidthRecord(stream),
                DefaultRowHeightRecord.sid => new DefaultRowHeightRecord(stream),
                DeltaRecord.sid => new DeltaRecord(stream),
                DimensionsRecord.sid => new DimensionsRecord(stream),
                DrawingGroupRecord.sid => new DrawingGroupRecord(stream),
                DrawingRecord.sid => new DrawingRecord(stream),
                DrawingSelectionRecord.sid => new DrawingSelectionRecord(stream),
                DSFRecord.sid => new DSFRecord(stream),
                DVALRecord.sid => new DVALRecord(stream),
                DVRecord.sid => new DVRecord(stream),
                EOFRecord.sid => new EOFRecord(stream),
                ExtendedFormatRecord.sid => new ExtendedFormatRecord(stream),
                ExternalNameRecord.sid => new ExternalNameRecord(stream),
                ExternSheetRecord.sid => new ExternSheetRecord(stream),
                ExtSSTRecord.sid => new ExtSSTRecord(stream),
                FilePassRecord.sid => new FilePassRecord(stream),
                FileSharingRecord.sid => new FileSharingRecord(stream),
                FnGroupCountRecord.sid => new FnGroupCountRecord(stream),
                FontRecord.sid => new FontRecord(stream),
                FooterRecord.sid => new FooterRecord(stream),
                FormatRecord.sid => new FormatRecord(stream),
                FormulaRecord.sid => new FormulaRecord(stream),
                GridsetRecord.sid => new GridsetRecord(stream),
                GutsRecord.sid => new GutsRecord(stream),
                HCenterRecord.sid => new HCenterRecord(stream),
                HeaderRecord.sid => new HeaderRecord(stream),
                HeaderFooterRecord.sid => new HeaderFooterRecord(stream),
                HideObjRecord.sid => new HideObjRecord(stream),
                HorizontalPageBreakRecord.sid => new HorizontalPageBreakRecord(stream),
                HyperlinkRecord.sid => new HyperlinkRecord(stream),
                IndexRecord.sid => new IndexRecord(stream),
                InterfaceEndRecord.sid => InterfaceEndRecord.Create(stream),
                InterfaceHdrRecord.sid => new InterfaceHdrRecord(stream),
                IterationRecord.sid => new IterationRecord(stream),
                LabelRecord.sid => new LabelRecord(stream),
                LabelSSTRecord.sid => new LabelSSTRecord(stream),
                LeftMarginRecord.sid => new LeftMarginRecord(stream),
                MergeCellsRecord.sid => new MergeCellsRecord(stream),
                MMSRecord.sid => new MMSRecord(stream),
                MulBlankRecord.sid => new MulBlankRecord(stream),
                MulRKRecord.sid => new MulRKRecord(stream),
                NameRecord.sid => new NameRecord(stream),
                NameCommentRecord.sid => new NameCommentRecord(stream),
                NoteRecord.sid => new NoteRecord(stream),
                NumberRecord.sid => new NumberRecord(stream),
                ObjectProtectRecord.sid => new ObjectProtectRecord(stream),
                ObjRecord.sid => new ObjRecord(stream),
                PaletteRecord.sid => new PaletteRecord(stream),
                PaneRecord.sid => new PaneRecord(stream),
                PasswordRecord.sid => new PasswordRecord(stream),
                PasswordRev4Record.sid => new PasswordRev4Record(stream),
                PrecisionRecord.sid => new PrecisionRecord(stream),
                PrintGridlinesRecord.sid => new PrintGridlinesRecord(stream),
                PrintHeadersRecord.sid => new PrintHeadersRecord(stream),
                PrintSetupRecord.sid => new PrintSetupRecord(stream),
                PrintSizeRecord.sid => new PrintSizeRecord(stream),
                ProtectionRev4Record.sid => new ProtectionRev4Record(stream),
                ProtectRecord.sid => new ProtectRecord(stream),
                RecalcIdRecord.sid => new RecalcIdRecord(stream),
                RefModeRecord.sid => new RefModeRecord(stream),
                RefreshAllRecord.sid => new RefreshAllRecord(stream),
                RightMarginRecord.sid => new RightMarginRecord(stream),
                RKRecord.sid => new RKRecord(stream),
                RowRecord.sid => new RowRecord(stream),
                SaveRecalcRecord.sid => new SaveRecalcRecord(stream),
                ScenarioProtectRecord.sid => new ScenarioProtectRecord(stream),
                SCLRecord.sid => new SCLRecord(stream),
                SelectionRecord.sid => new SelectionRecord(stream),
                SeriesRecord.sid => new SeriesRecord(stream),
                SeriesTextRecord.sid => new SeriesTextRecord(stream),
                SharedFormulaRecord.sid => new SharedFormulaRecord(stream),
                SSTRecord.sid => new SSTRecord(stream),
                StringRecord.sid => new StringRecord(stream),
                StyleRecord.sid => new StyleRecord(stream),
                SupBookRecord.sid => new SupBookRecord(stream),
                TabIdRecord.sid => new TabIdRecord(stream),
                TableRecord.sid => new TableRecord(stream),
                TableStylesRecord.sid => new TableStylesRecord(stream),
                TextObjectRecord.sid => new TextObjectRecord(stream),
                TopMarginRecord.sid => new TopMarginRecord(stream),
                UncalcedRecord.sid => new UncalcedRecord(stream),
                UseSelFSRecord.sid => new UseSelFSRecord(stream),
                UserSViewBegin.sid => new UserSViewBegin(stream),
                UserSViewEnd.sid => new UserSViewEnd(stream),
                ValueRangeRecord.sid => new ValueRangeRecord(stream),
                VCenterRecord.sid => new VCenterRecord(stream),
                VerticalPageBreakRecord.sid => new VerticalPageBreakRecord(stream),
                WindowOneRecord.sid => new WindowOneRecord(stream),
                WindowProtectRecord.sid => new WindowProtectRecord(stream),
                WindowTwoRecord.sid => new WindowTwoRecord(stream),
                WriteAccessRecord.sid => new WriteAccessRecord(stream),
                WriteProtectRecord.sid => new WriteProtectRecord(stream),
                WSBoolRecord.sid => new WSBoolRecord(stream),
                SheetExtRecord.sid => new SheetExtRecord(stream),
                AreaFormatRecord.sid => new AreaFormatRecord(stream),
                AreaRecord.sid => new AreaRecord(stream),
                AttachedLabelRecord.sid => new AttachedLabelRecord(stream),
                AxcExtRecord.sid => new AxcExtRecord(stream),
                AxisLineFormatRecord.sid => new AxisLineFormatRecord(stream),
                AxisParentRecord.sid => new AxisParentRecord(stream),
                AxisRecord.sid => new AxisRecord(stream),
                AxesUsedRecord.sid => new AxesUsedRecord(stream),
                BarRecord.sid => new BarRecord(stream),
                BeginRecord.sid => new BeginRecord(stream),
                BopPopCustomRecord.sid => new BopPopCustomRecord(stream),
                BopPopRecord.sid => new BopPopRecord(stream),
                CatLabRecord.sid => new CatLabRecord(stream),
                CatSerRangeRecord.sid => new CatSerRangeRecord(stream),
                Chart3DBarShapeRecord.sid => new Chart3DBarShapeRecord(stream),
                Chart3dRecord.sid => new Chart3dRecord(stream),
                ChartEndObjectRecord.sid => new ChartEndObjectRecord(stream),
                ChartFormatRecord.sid => new ChartFormatRecord(stream),
                ChartFRTInfoRecord.sid => new ChartFRTInfoRecord(stream),
                ChartStartObjectRecord.sid => new ChartStartObjectRecord(stream),
                CrtLayout12ARecord.sid => new CrtLayout12ARecord(stream),
                CrtLayout12Record.sid => new CrtLayout12Record(stream),
                CrtLineRecord.sid => new CrtLineRecord(stream),
                CrtLinkRecord.sid => new CrtLinkRecord(stream),
                CrtMlFrtContinueRecord.sid => new CrtMlFrtContinueRecord(stream),
                CrtMlFrtRecord.sid => new CrtMlFrtRecord(stream),
                DataFormatRecord.sid => new DataFormatRecord(stream),
                DataLabExtContentsRecord.sid => new DataLabExtContentsRecord(stream),
                DataLabExtRecord.sid => new DataLabExtRecord(stream),
                DatRecord.sid => new DatRecord(stream),
                DefaultTextRecord.sid => new DefaultTextRecord(stream),
                DropBarRecord.sid => new DropBarRecord(stream),
                EndBlockRecord.sid => new EndBlockRecord(stream),
                EndRecord.sid => new EndRecord(stream),
                LinkedDataRecord.sid => new LinkedDataRecord(stream),
                Fbi2Record.sid => new Fbi2Record(stream),
                FbiRecord.sid => new FbiRecord(stream),
                FontIndexRecord.sid => new FontIndexRecord(stream),
                FrameRecord.sid => new FrameRecord(stream),
                FrtFontListRecord.sid => new FrtFontListRecord(stream),
                GelFrameRecord.sid => new GelFrameRecord(stream),
                IFmtRecordRecord.sid => new IFmtRecordRecord(stream),
                LegendExceptionRecord.sid => new LegendExceptionRecord(stream),
                LegendRecord.sid => new LegendRecord(stream),
                LineFormatRecord.sid => new LineFormatRecord(stream),
                MarkerFormatRecord.sid => new MarkerFormatRecord(stream),
                ObjectLinkRecord.sid => new ObjectLinkRecord(stream),
                PicFRecord.sid => new PicFRecord(stream),
                PieFormatRecord.sid => new PieFormatRecord(stream),
                PieRecord.sid => new PieRecord(stream),
                PlotAreaRecord.sid => new PlotAreaRecord(stream),
                PlotGrowthRecord.sid => new PlotGrowthRecord(stream),
                PosRecord.sid => new PosRecord(stream),
                RadarAreaRecord.sid => new RadarAreaRecord(stream),
                RadarRecord.sid => new RadarRecord(stream),
                RichTextStreamRecord.sid => new RichTextStreamRecord(stream),
                ScatterRecord.sid => new ScatterRecord(stream),
                SerAuxErrBarRecord.sid => new SerAuxErrBarRecord(stream),
                SerAuxTrendRecord.sid => new SerAuxTrendRecord(stream),
                SerFmtRecord.sid => new SerFmtRecord(stream),
                SeriesIndexRecord.sid => new SeriesIndexRecord(stream),
                SeriesListRecord.sid => new SeriesListRecord(stream),
                SerParentRecord.sid => new SerParentRecord(stream),
                SerToCrtRecord.sid => new SerToCrtRecord(stream),
                ShapePropsStreamRecord.sid => new ShapePropsStreamRecord(stream),
                ShtPropsRecord.sid => new ShtPropsRecord(stream),
                StartBlockRecord.sid => new StartBlockRecord(stream),
                SurfRecord.sid => new SurfRecord(stream),
                TextPropsStreamRecord.sid => new TextPropsStreamRecord(stream),
                TextRecord.sid => new TextRecord(stream),
                TickRecord.sid => new TickRecord(stream),
                UnitsRecord.sid => new UnitsRecord(stream),
                YMultRecord.sid => new YMultRecord(stream),
                DataItemRecord.sid => new DataItemRecord(stream),
                ExtendedPivotTableViewFieldsRecord.sid => new ExtendedPivotTableViewFieldsRecord(stream),
                PageItemRecord.sid => new PageItemRecord(stream),
                StreamIDRecord.sid => new StreamIDRecord(stream),
                ViewDefinitionRecord.sid => new ViewDefinitionRecord(stream),
                ViewFieldsRecord.sid => new ViewFieldsRecord(stream),
                ViewSourceRecord.sid => new ViewSourceRecord(stream),
                AutoFilterRecord.sid => new AutoFilterRecord(stream),
                FilterModeRecord.sid => new FilterModeRecord(stream),
                Excel9FileRecord.sid => new Excel9FileRecord(stream),
                _ => new UnknownRecord(stream)
            };
        }
        /**
	     * cache of the recordsToMap();
	     */
        //private static Dictionary<short, I_RecordCreator> _recordCreatorsById = null;//RecordsToMap(recordClasses);

        private static short[] _allKnownRecordSIDs;
        /**
         * Debug / diagnosis method<br/>
         * Gets the POI implementation class for a given <tt>sid</tt>.  Only a subset of the any BIFF
         * records are actually interpreted by POI.  A few others are known but not interpreted
         * (see {@link UnknownRecord#getBiffName(int)}).
         * @return the POI implementation class for the specified record <tt>sid</tt>.
         * <code>null</code> if the specified record is not interpreted by POI.
         */
        public static Type GetRecordClass(int sid)
        {
            return _recordTypes.TryGetValue((short)sid, out var type) ? type : null;
        }
        /**
         * Changes the default capacity (10000) to handle larger files
         */

        public static void SetCapacity(int capacity)
        {
            NUM_RECORDS = capacity;
        }

        /**
         * Create an array of records from an input stream
         *
         * @param in the InputStream from which the records will be
         *           obtained
         *
         * @return an array of Records Created from the InputStream
         *
         * @exception RecordFormatException on error Processing the
         *            InputStream
         */

        public static List<Record> CreateRecords(Stream in1)
        {
            List<Record> records = new List<Record>(NUM_RECORDS);


            RecordFactoryInputStream recStream = new RecordFactoryInputStream(in1, true);

            Record record;
            while ((record = recStream.NextRecord()) != null)
            {
                records.Add(record);
            }

            return records;
        }

        public static Record[] CreateRecord(RecordInputStream in1)
        {
            Record record = CreateSingleRecord(in1);
            if (record is DBCellRecord)
            {
                // Not needed by POI.  Regenerated from scratch by POI when spreadsheet is written
                return new Record[] { null, };
            }
            if (record is RKRecord)
            {
                return new Record[] { ConvertToNumberRecord((RKRecord)record), };
            }
            if (record is MulRKRecord)
            {
                return ConvertRKRecords((MulRKRecord)record);
            }
            return new Record[] { record, };
        }
        /**
         * Converts a {@link MulBlankRecord} into an equivalent array of {@link BlankRecord}s
         */
        public static BlankRecord[] ConvertBlankRecords(MulBlankRecord mbk)
        {
            BlankRecord[] mulRecs = new BlankRecord[mbk.NumColumns];
            for (int k = 0; k < mbk.NumColumns; k++)
            {
                BlankRecord br = new BlankRecord();

                br.Column = k + mbk.FirstColumn;
                br.Row = mbk.Row;
                br.XFIndex = mbk.GetXFAt(k);
                mulRecs[k] = br;
            }
            return mulRecs;
        }

        public static Record CreateSingleRecord(RecordInputStream in1)
        {
            return CreateBySid(in1.Sid, in1);
        }
        /// <summary>
        /// RK record is a slightly smaller alternative to NumberRecord
        /// POI likes NumberRecord better
        /// </summary>
        /// <param name="rk">The rk.</param>
        /// <returns></returns>
        public static NumberRecord ConvertToNumberRecord(RKRecord rk)
        {
            NumberRecord num = new NumberRecord();

            num.Column = (rk.Column);
            num.Row = (rk.Row);
            num.XFIndex = (rk.XFIndex);
            num.Value = (rk.RKNumber);
            return num;
        }
        /// <summary>
        /// Converts a MulRKRecord into an equivalent array of NumberRecords
        /// </summary>
        /// <param name="mrk">The MRK.</param>
        /// <returns></returns>
        public static NumberRecord[] ConvertRKRecords(MulRKRecord mrk)
        {

            NumberRecord[] mulRecs = new NumberRecord[mrk.NumColumns];
            for (int k = 0; k < mrk.NumColumns; k++)
            {
                NumberRecord nr = new NumberRecord();

                nr.Column = ((short)(k + mrk.FirstColumn));
                nr.Row = (mrk.Row);
                nr.XFIndex = (mrk.GetXFAt(k));
                nr.Value = (mrk.GetRKNumberAt(k));
                mulRecs[k] = nr;
            }
            return mulRecs;
        }
        public static short[] GetAllKnownRecordSIDs()
        {
            if (_allKnownRecordSIDs == null)
            {
                _allKnownRecordSIDs = _recordTypes.Keys.ToArray();
                Array.Sort(_allKnownRecordSIDs);
            }

            return (short[])_allKnownRecordSIDs.Clone();
        }
    }
}

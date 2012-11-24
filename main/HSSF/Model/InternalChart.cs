using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Chart;
using NPOI.HSSF.Record.Aggregates;
using NPOI.Util;
using NPOI.SS.UserModel;

namespace NPOI.HSSF.Model
{

    [Serializable]
    public class InternalChart
    {
        private HeaderRecord header;
        private FooterRecord footer;
        private HCenterRecord _hCenter;
        private VCenterRecord _vCenter;
        private LeftMarginRecord _leftMargin;
        private RightMarginRecord _rightMargin;
        private TopMarginRecord _topMargin;
        private BottomMarginRecord _bottomMargin;
        // fix warning CS0169 "never used": private Record _pls;
        private PrintSetupRecord printSetup;

        private HeaderFooterRecord _headerFooter;
        private PrintSizeRecord _printSize;
        private List<PLSAggregate> _plsRecords;
        private List<HeaderFooterRecord> _sviewHeaderFooters = new List<HeaderFooterRecord>();

        private ProtectRecord _protect;
        private ChartFRTInfoRecord _chartFrtInfo;
        protected List<RecordBase> records = null;
        public InternalChart(RecordStream rs)
        {
            _plsRecords = new List<PLSAggregate>();
            records = new List<RecordBase>(128);

            if (rs.PeekNextSid() != BOFRecord.sid)
            {
                throw new Exception("BOF record expected");
            }
            BOFRecord bof = (BOFRecord)rs.GetNext();
            if (bof.Type != BOFRecord.TYPE_CHART)
            {
                // TODO - fix junit tests throw new RuntimeException("Bad BOF record type");
                throw new RuntimeException("Bad BOF record type");
            }

            records.Add(bof);
            while (rs.HasNext())
            {
                int recSid = rs.PeekNextSid();

                Record.Record rec = rs.GetNext();
                if (recSid == EOFRecord.sid)
                {
                    records.Add(rec);
                    break;
                }

                if (recSid == ChartRecord.sid)
                {

                    continue;
                }

                if (recSid == ChartFRTInfoRecord.sid)
                {
                    _chartFrtInfo = (ChartFRTInfoRecord)rec;
                }
                else if (recSid == HeaderRecord.sid)
                {
                    header = (HeaderRecord)rec;
                }
                else if (recSid == FooterRecord.sid)
                {
                    footer = (FooterRecord)rec;
                }
                else if (recSid == HCenterRecord.sid)
                {
                    _hCenter = (HCenterRecord)rec;
                }
                else if (recSid == VCenterRecord.sid)
                {
                    _vCenter = (VCenterRecord)rec;
                }
                else if (recSid == LeftMarginRecord.sid)
                {
                    _leftMargin = (LeftMarginRecord)rec;
                }
                else if (recSid == RightMarginRecord.sid)
                {
                    _rightMargin = (RightMarginRecord)rec;
                }
                else if (recSid == TopMarginRecord.sid)
                {
                    _topMargin = (TopMarginRecord)rec;
                }
                else if (recSid == BottomMarginRecord.sid)
                {
                    _bottomMargin = (BottomMarginRecord)rec;
                }
                else if (recSid == UnknownRecord.PLS_004D) // PLS
                {
                    PLSAggregate pls = new PLSAggregate(rs);
                    PLSAggregateVisitor rv = new PLSAggregateVisitor(records);
                    pls.VisitContainedRecords(rv);
                    _plsRecords.Add(pls);

                    continue;
                }
                else if (recSid == PrintSetupRecord.sid)
                {
                    printSetup = (PrintSetupRecord)rec;
                }
                else if (recSid == PrintSizeRecord.sid)
                {
                    _printSize = (PrintSizeRecord)rec;
                }
                else if (recSid == HeaderFooterRecord.sid)
                {
                    HeaderFooterRecord hf = (HeaderFooterRecord)rec;
                    if (hf.IsCurrentSheet)
                        _headerFooter = hf;
                    else
                        _sviewHeaderFooters.Add(hf);
                }
                else if (recSid == ProtectRecord.sid)
                {
                    _protect = (ProtectRecord)rec;
                }
                records.Add(rec);
            }
            
        }
        private class PLSAggregateVisitor : RecordVisitor
        {
            private List<RecordBase> container;
            public PLSAggregateVisitor(List<RecordBase> container)
            {
                this.container = container;
            }
            #region RecordVisitor 成员

            public void VisitRecord(NPOI.HSSF.Record.Record r)
            {
                container.Add((RecordBase)r);
            }

            #endregion
        }
        private void CheckNotPresent(Record.Record rec)
        {
            if (rec != null)
            {
                throw new RecordFormatException("Duplicate PageSettingsBlock record (sid=0x"
                        + StringUtil.ToHexString(rec.Sid) + ")");
            }
        }
        private InternalChart()
        {
        }
        List<Record.Record> recores = null;

        private IMargin GetMarginRec(MarginType margin)
        {
            switch (margin)
            {
                case MarginType.LeftMargin: return _leftMargin;
                case MarginType.RightMargin: return _rightMargin;
                case MarginType.TopMargin: return _topMargin;
                case MarginType.BottomMargin: return _bottomMargin;
                default:
                    throw new InvalidOperationException("Unknown margin constant:  " + (short)margin);
            }
        }


        /**
         * Gets the size of the margin in inches.
         * @param margin which margin to Get
         * @return the size of the margin
         */
        public double GetMargin(MarginType margin)
        {
            IMargin m = GetMarginRec(margin);
            if (m != null)
            {
                return m.Margin;
            }
            else
            {
                switch (margin)
                {
                    case MarginType.LeftMargin:
                        return .7;
                    case MarginType.RightMargin:
                        return .7;
                    case MarginType.TopMargin:
                        return .75;
                    case MarginType.BottomMargin:
                        return .75;
                }
                throw new InvalidOperationException("Unknown margin constant:  " + margin);
            }
        }

        /**
         * Sets the size of the margin in inches.
         * @param margin which margin to Get
         * @param size the size of the margin
         */
        public void SetMargin(MarginType margin, double size)
        {
            IMargin m = GetMarginRec(margin);
            if (m == null)
            {
                switch (margin)
                {
                    case MarginType.LeftMargin:
                        _leftMargin = new LeftMarginRecord();
                        m = _leftMargin;
                        break;
                    case MarginType.RightMargin:
                        _rightMargin = new RightMarginRecord();
                        m = _rightMargin;
                        break;
                    case MarginType.TopMargin:
                        _topMargin = new TopMarginRecord();
                        m = _topMargin;
                        break;
                    case MarginType.BottomMargin:
                        _bottomMargin = new BottomMarginRecord();
                        m = _bottomMargin;
                        break;
                    default:
                        throw new InvalidOperationException("Unknown margin constant:  " + margin);
                }
            }
            m.Margin = size;
        }

        #region Creation Method for creating a new chart

        private static InternalChart CreateChartSheet()
        {
            InternalChart retval = new InternalChart();
            List<Record.Record> records = new List<Record.Record>(30);
            retval.recores = records;

            records.Add(CreateBOFRecord());
            records.Add(CreateChartFRTInfoRecord());
            records.Add(new HeaderRecord(string.Empty));
            records.Add(new FooterRecord(string.Empty));
            records.Add(CreateHCenterRecord());
            records.Add(CreateVCenterRecord());
            records.Add((LeftMarginRecord)CreateMarginRecord(MarginType.LeftMargin, 0.7));
            records.Add((RightMarginRecord)CreateMarginRecord(MarginType.RightMargin, 0.7));
            records.Add((TopMarginRecord)CreateMarginRecord(MarginType.TopMargin, 0.7));
            records.Add((BottomMarginRecord)CreateMarginRecord(MarginType.BottomMargin, 0.7));
            //ignore pls
            
            records.Add(CreatePrintSetupRecord());
            records.Add(CreatePrintSizeRecord());
            records.Add(CreateFontBasisRecord1());
            records.Add(CreateFontBasisRecord2());
            //records.Add(CreateHeaderFooterRecord()); //ignore this record
            records.Add(new ProtectRecord(false));
            //records.Add(CreateDrawingRecord());
            records.Add(new UnitsRecord());

            //CHARTFOMATS = Chart Begin *2FONTLIST Scl PlotGrowth [FRAME] *SERIESFORMAT *SS 
            //ShtProps *2DFTTEXT AxesUsed 1*2AXISPARENT [CrtLayout12A] [DAT] *ATTACHEDLABEL 
            //[CRTMLFRT] *([DataLabExt StartObject] ATTACHEDLABEL [EndObject]) [TEXTPROPS] *2CRTMLFRT End
            records.Add(CreateChartRecord(0, 0, 32341968, 14745600));
            records.Add(new BeginRecord());
            records.Add(CreateSCLRecord(1, 1));
            records.Add(CreatePlotGrowthRecord(65536, 65536));
            //Frame
            records.Add(CreateFrameRecord1());
            records.Add(new BeginRecord());
            records.Add(CreateLineFormatRecord(true));
            records.Add(CreateAreaFormatRecord1());
            records.Add(new EndRecord());

            //an empty chart has no SERIESFORMAT
            //CreateSERIESFORMAT(records);
            //an empty chart has no SS
            //CreateSS(records);

            records.Add(CreateShtPropsRecord());

            //add 
            records.Add(new EOFRecord());
            return retval;
        }
        private static ShtPropsRecord CreateShtPropsRecord()
        {
            ShtPropsRecord r = new ShtPropsRecord();
            r.Flags = 2;  //set bit fPlotVisOnly 1
            return r;
        }

        #region SERIESFORMAT

        //SERIESFORMAT = Series Begin 4AI *SS (SerToCrt / (SerParent (SerAuxTrend / SerAuxErrBar))) 
        //*(LegendException [Begin ATTACHEDLABEL [TEXTPROPS] End]) End
        private static void CreateSERIESFORMAT(List<Record.Record> records)
        {
            
            //Series
            records.Add(CreateSeriesRecord());
            records.Add(new BeginRecord());
            //  4AI

            //  DataFormat
            records.Add(CreateDataFormatRecord());
            records.Add(new BeginRecord());
            records.Add(CreateChart3DBarShapeRecord());
            records.Add(new EndRecord()); //end dataformat

            records.Add(new EndRecord());  //end series
        }
        private static SeriesRecord CreateSeriesRecord()
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

        private static DataFormatRecord CreateDataFormatRecord()
        {
            DataFormatRecord r = new DataFormatRecord();
            r.PointNumber = ((short)-1);
            r.SeriesIndex = ((short)0);
            r.SeriesNumber = ((short)0);
            r.UseExcel4Colors = (false);
            return r;
        }
        #region SS
        //SS = DataFormat Begin [Chart3DBarShape] [LineFormat AreaFormat PieFormat] [SerFmt] 
        //[GELFRAME] [MarkerFormat] [AttachedLabel] *2SHAPEPROPS [CRTMLFRT] End
        private static void CreateSS(List<NPOI.HSSF.Record.Record> records)
        {
            records.Add(CreateDataFormatRecord());
            records.Add(new BeginRecord());
            //Chart3DBarShape
            records.Add(CreateChart3DBarShapeRecord());

            records.Add(new EndRecord());
        }

        private static Chart3DBarShapeRecord CreateChart3DBarShapeRecord()
        {
            Chart3DBarShapeRecord r = new Chart3DBarShapeRecord();
            r.Riser = 0;
            r.Taper = 0;
            return r;
        }
        #endregion
        #endregion
        
        
        private static DrawingRecord CreateDrawingRecord()
        {
            //throw new NotImplementedException();
            byte[] drawingData = HexRead.ReadFromString("0F 00 02 F0 48 00 00 00 30 00 08 F0 " +
                                            "08 00 00 00 01 00 00 00 00 0C 00 00 0F 00 03 F0 " +
                                            "30 00 00 00 0F 00 04 F0 28 00 00 00 01 00 09 F0 " +
                                            "10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                                            "00 00 00 00 02 00 0A F0 08 00 00 00 00 0C 00 00 " +
                                            "05 00 00 00");

            DrawingRecord retval = new DrawingRecord();
            retval.Data = drawingData;
            return retval;
        }

        private static HeaderFooterRecord CreateHeaderFooterRecord()
        {
            throw new NotImplementedException();
        }
        private static IMargin CreateMarginRecord(MarginType margin, double size)
        {
            IMargin m;
            switch (margin)
            {
                case MarginType.LeftMargin:
                    m = new LeftMarginRecord();
                    break;
                case MarginType.RightMargin:
                    m = new RightMarginRecord();
                    break;
                case MarginType.TopMargin:
                    m = new TopMarginRecord();
                    break;
                case MarginType.BottomMargin:
                    m = new BottomMarginRecord();
                    break;
                default:
                    throw new InvalidOperationException("Unknown margin constant:  " + margin);
            }
            m.Margin = size;
            return m;
        }

        private static PrintSetupRecord CreatePrintSetupRecord()
        {
            PrintSetupRecord retval = new PrintSetupRecord();

            retval.PaperSize = ((short)0);
            retval.Scale = ((short)18);
            retval.PageStart = ((short)1);
            retval.FitWidth = ((short)1);
            retval.FitHeight = ((short)1);
            retval.Options = ((short)4);
            retval.HResolution = ((short)0);
            retval.VResolution = ((short)0);
            retval.HeaderMargin = (0.3);
            retval.FooterMargin = (0.3);
            retval.Copies = ((short)1);
            return retval;
        }
        private static BOFRecord CreateBOFRecord()
        {
            BOFRecord retval = new BOFRecord();
            retval.Version = ((short)600);
            retval.Type = BOFRecord.TYPE_CHART;
            retval.Build = ((short)0x1CFE);
            retval.BuildYear = ((short)1997);
            retval.HistoryBitMask = (0x40C9);
            retval.RequiredVersion = (106);
            return retval;
        }
        private static PrintSizeRecord CreatePrintSizeRecord()
        {
            PrintSizeRecord retval = new PrintSizeRecord();
            retval.PrintSize = 3;
            return retval;
        }
        private static ChartFRTInfoRecord CreateChartFRTInfoRecord()
        {
            ChartFRTInfoRecord retval = new ChartFRTInfoRecord();
            return retval;
        }

        private static FooterRecord CreateFooterRecord()
        {
            FooterRecord retval = new FooterRecord(string.Empty);
            return retval;
        }
        private static HCenterRecord CreateHCenterRecord()
        {
            HCenterRecord r = new HCenterRecord();
            r.HCenter = (false);
            return r;
        }

        private static VCenterRecord CreateVCenterRecord()
        {
            VCenterRecord r = new VCenterRecord();
            r.VCenter = (false);
            return r;
        }

        
        private static FontBasisRecord CreateFontBasisRecord1()
        {
            FontBasisRecord r = new FontBasisRecord();
            r.XBasis = ((short)9720);
            r.YBasis = ((short)4350);
            r.HeightBasis = ((short)240);
            r.Scale = ((short)0);
            r.IndexToFontTable = ((short)24);
            return r;
        }

        private static FontBasisRecord CreateFontBasisRecord2()
        {
            FontBasisRecord r = CreateFontBasisRecord1();
            r.Scale = 1;
            r.IndexToFontTable = ((short)25);
            return r;
        }

        //CHARTFORMATS

        private static ChartRecord CreateChartRecord(int x, int y, int width, int height)
        {
            ChartRecord r = new ChartRecord();
            r.X = (x);
            r.Y = (y);
            r.Width = (width);
            r.Height = (height);
            return r;
        }

        private static PlotGrowthRecord CreatePlotGrowthRecord(int horizScale, int vertScale)
        {
            PlotGrowthRecord r = new PlotGrowthRecord();
            r.HorizontalScale = (horizScale);
            r.VerticalScale = (vertScale);
            return r;
        }

        private static SCLRecord CreateSCLRecord(short numerator, short denominator)
        {
            SCLRecord r = new SCLRecord();
            r.Denominator = (denominator);
            r.Numerator = (numerator);
            return r;
        }

        private static FrameRecord CreateFrameRecord1()
        {
            FrameRecord r = new FrameRecord();
            r.BorderType = (FrameRecord.BORDER_TYPE_REGULAR);
            r.IsAutoSize = (false);
            r.IsAutoPosition = (true);
            return r;
        }

        private static FrameRecord CreateFrameRecord2()
        {
            FrameRecord r = new FrameRecord();
            r.BorderType = (FrameRecord.BORDER_TYPE_REGULAR);
            r.IsAutoSize = (true);
            r.IsAutoPosition = (true);
            return r;
        }

        private static AreaFormatRecord CreateAreaFormatRecord1()
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

        private static AreaFormatRecord CreateAreaFormatRecord2()
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

        private static LineFormatRecord CreateLineFormatRecord(bool drawTicks)
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

        private static LineFormatRecord CreateLineFormatRecord2()
        {
            LineFormatRecord r = new LineFormatRecord();
            r.LineColor = (0x00808080);
            r.LinePattern = LineFormatRecord.LINE_PATTERN_SOLID;
            r.Weight = ((short)0);
            r.IsAuto = (false);
            r.IsDrawTicks = (false);
            r.IsUnknown = (false);
            r.ColourPaletteIndex = ((short)23);
            return r;
        }
        #endregion
    }
}

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
    public class ChartSheet
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
        public ChartSheet(RecordStream rs)
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
                    _plsRecords.Add(new PLSAggregate(rs));
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

        private void CheckNotPresent(Record.Record rec)
        {
            if (rec != null)
            {
                throw new RecordFormatException("Duplicate PageSettingsBlock record (sid=0x"
                        + StringUtil.ToHexString(rec.Sid) + ")");
            }
        }
        private ChartSheet()
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

        private static ChartSheet CreateChartSheet()
        {
            ChartSheet retval = new ChartSheet();
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
            records.Add(CreatePrintSetupRecord());
            records.Add(CreatePrintSizeRecord());
            //records.Add(CreateHeaderFooterRecord()); //ignore this record
            records.Add(new ProtectRecord(false));
            records.Add(CreateDrawingRecord());
            records.Add(new UnitsRecord());

            return retval;
        }

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

        private class CHARTFOMATS
        {
            public CHARTFOMATS(RecordStream rs)
            {
            }
        }
    }
}

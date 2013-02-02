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

namespace NPOI.HSSF.Record.Aggregates
{

    using System;
    using System.Collections;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using System.Collections.Generic;
    /**
     * Groups the page settings records for a worksheet.<p/>
     * 
     * See OOO excelfileformat.pdf sec 4.4 'Page Settings Block'
     * 
     * @author Josh Micich
     */
    public class PageSettingsBlock : RecordAggregate
    {
        // Every one of these component records is optional 
        // (The whole PageSettingsBlock may not be present) 
        private PageBreakRecord _rowBreaksRecord;
        private PageBreakRecord _columnBreaksRecord;
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
        private Record _bitmap;
        private HeaderFooterRecord _headerFooter;

        private List<HeaderFooterRecord> _sviewHeaderFooters = new List<HeaderFooterRecord>();
        private List<PLSAggregate> _plsRecords;

        private Record _printSize;

        public PageSettingsBlock(RecordStream rs)
        {
            _plsRecords = new List<PLSAggregate>();
            while (ReadARecord(rs)) ;
        }

        /**
         * Creates a PageSettingsBlock with default settings
         */
        public PageSettingsBlock()
        {
            _plsRecords = new List<PLSAggregate>();
            _rowBreaksRecord = new HorizontalPageBreakRecord();
            _columnBreaksRecord = new VerticalPageBreakRecord();
            header = new HeaderRecord(string.Empty);
            footer = new FooterRecord(string.Empty);
            _hCenter = CreateHCenter();
            _vCenter = CreateVCenter();
            printSetup = CreatePrintSetup();
        }
        /**
         * @return <c>true</c> if the specified Record sid is one belonging to the 
         * 'Page Settings Block'.
         */
        public static bool IsComponentRecord(int sid)
        {
            switch (sid)
            {
                case HorizontalPageBreakRecord.sid:
                case VerticalPageBreakRecord.sid:
                case HeaderRecord.sid:
                case FooterRecord.sid:
                case HCenterRecord.sid:
                case VCenterRecord.sid:
                case LeftMarginRecord.sid:
                case RightMarginRecord.sid:
                case TopMarginRecord.sid:
                case BottomMarginRecord.sid:
                case UnknownRecord.PLS_004D:
                case PrintSetupRecord.sid:
                case UnknownRecord.BITMAP_00E9:
                //case UnknownRecord.PRINTSIZE_0033:
                case PrintSizeRecord.sid:
                case HeaderFooterRecord.sid: // extra header/footer settings supported by Excel 2007
                    return true;
            }
            return false;
        }

        private bool ReadARecord(RecordStream rs)
        {
            switch (rs.PeekNextSid())
            {
                case HorizontalPageBreakRecord.sid:
                    CheckNotPresent(_rowBreaksRecord);
                    _rowBreaksRecord = (PageBreakRecord)rs.GetNext();
                    break;
                case VerticalPageBreakRecord.sid:
                    CheckNotPresent(_columnBreaksRecord);
                    _columnBreaksRecord = (PageBreakRecord)rs.GetNext();
                    break;
                case HeaderRecord.sid:
                    CheckNotPresent(header);
                    header = (HeaderRecord)rs.GetNext();
                    break;
                case FooterRecord.sid:
                    CheckNotPresent(footer);
                    footer = (FooterRecord)rs.GetNext();
                    break;
                case HCenterRecord.sid:
                    CheckNotPresent(_hCenter);
                    _hCenter = (HCenterRecord)rs.GetNext();
                    break;
                case VCenterRecord.sid:
                    CheckNotPresent(_vCenter);
                    _vCenter = (VCenterRecord)rs.GetNext();
                    break;
                case LeftMarginRecord.sid:
                    CheckNotPresent(_leftMargin);
                    _leftMargin = (LeftMarginRecord)rs.GetNext();
                    break;
                case RightMarginRecord.sid:
                    CheckNotPresent(_rightMargin);
                    _rightMargin = (RightMarginRecord)rs.GetNext();
                    break;
                case TopMarginRecord.sid:
                    CheckNotPresent(_topMargin);
                    _topMargin = (TopMarginRecord)rs.GetNext();
                    break;
                case BottomMarginRecord.sid:
                    CheckNotPresent(_bottomMargin);
                    _bottomMargin = (BottomMarginRecord)rs.GetNext();
                    break;
                case UnknownRecord.PLS_004D: // PLS
                    _plsRecords.Add(new PLSAggregate(rs));
                    break;
                case PrintSetupRecord.sid:
                    CheckNotPresent(printSetup);
                    printSetup = (PrintSetupRecord)rs.GetNext();
                    break;
                case UnknownRecord.BITMAP_00E9: // BITMAP
                    CheckNotPresent(_bitmap);
                    _bitmap = rs.GetNext();
                    break;
                case PrintSizeRecord.sid:
                    CheckNotPresent(_printSize);
                    _printSize = rs.GetNext();
                    break;
                case HeaderFooterRecord.sid:
                    HeaderFooterRecord hf = (HeaderFooterRecord)rs.GetNext();
                    if (hf.IsCurrentSheet)
                        _headerFooter = hf;
                    else
                        _sviewHeaderFooters.Add(hf);
                    break;
                default:
                    // all other record types are not part of the PageSettingsBlock
                    return false;
            }
            return true;
        }
        private void CheckNotPresent(Record rec)
        {
            if (rec != null)
            {
                throw new RecordFormatException("Duplicate PageSettingsBlock record (sid=0x"
                        + StringUtil.ToHexString(rec.Sid) + ")");
            }
        }
        private PageBreakRecord RowBreaksRecord
        {
            get
            {
                if (_rowBreaksRecord == null)
                {
                    _rowBreaksRecord = new HorizontalPageBreakRecord();
                }
                return _rowBreaksRecord;
            }
        }

        private PageBreakRecord ColumnBreaksRecord
        {
            get
            {
                if (_columnBreaksRecord == null)
                {
                    _columnBreaksRecord = new VerticalPageBreakRecord();
                }
                return _columnBreaksRecord;
            }
        }

        public IEnumerator GetEnumerator()
        {
            return _plsRecords.GetEnumerator();
        }

        /**
         * Sets a page break at the indicated column
         *
         */
        public void SetColumnBreak(int column, int fromRow, int toRow)
        {
            this.ColumnBreaksRecord.AddBreak(column, fromRow, toRow);
        }

        /**
         * Removes a page break at the indicated column
         *
         */
        public void RemoveColumnBreak(int column)
        {
            this.ColumnBreaksRecord.RemoveBreak(column);
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            VisitIfPresent(_rowBreaksRecord, rv);
            VisitIfPresent(_columnBreaksRecord, rv);
            // Write out empty header / footer records if these are missing
            if (header == null)
            {
                rv.VisitRecord(new HeaderRecord(""));
            }
            else
            {
                rv.VisitRecord(header);
            }
            if (footer == null)
            {
                rv.VisitRecord(new FooterRecord(""));
            }
            else
            {
                rv.VisitRecord(footer);
            }
            VisitIfPresent(_hCenter, rv);
            VisitIfPresent(_vCenter, rv);
            VisitIfPresent(_leftMargin, rv);
            VisitIfPresent(_rightMargin, rv);
            VisitIfPresent(_topMargin, rv);
            VisitIfPresent(_bottomMargin, rv);
		    foreach (RecordAggregate pls in _plsRecords) {
			    pls.VisitContainedRecords(rv);
		    }
            VisitIfPresent(printSetup, rv);
            
            VisitIfPresent(_printSize, rv);
            VisitIfPresent(_headerFooter, rv);
            VisitIfPresent(_bitmap, rv);
        }
        private static void VisitIfPresent(Record r, RecordVisitor rv)
        {
            if (r != null)
            {
                rv.VisitRecord(r);
            }
        }
        private static void VisitIfPresent(PageBreakRecord r, RecordVisitor rv)
        {
            if (r != null)
            {
                if (r.IsEmpty)
                {
                    // its OK to not serialize empty page break records
                    return;
                }
                rv.VisitRecord(r);
            }
        }

        /**
         * Creates the HCenter Record and sets it to false (don't horizontally center)
         */
        private static HCenterRecord CreateHCenter()
        {
            HCenterRecord retval = new HCenterRecord();

            retval.HCenter = (false);
            return retval;
        }

        /**
         * Creates the VCenter Record and sets it to false (don't horizontally center)
        */
        private static VCenterRecord CreateVCenter()
        {
            VCenterRecord retval = new VCenterRecord();

            retval.VCenter = (false);
            return retval;
        }

        /**
         * Creates the PrintSetup Record and sets it to defaults and marks it invalid
         * @see org.apache.poi.hssf.record.PrintSetupRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a PrintSetupRecord
         */
        private static PrintSetupRecord CreatePrintSetup()
        {
            PrintSetupRecord retval = new PrintSetupRecord();

            retval.PaperSize = ((short)1);
            retval.Scale = ((short)100);
            retval.PageStart = ((short)1);
            retval.FitWidth = ((short)1);
            retval.FitHeight = ((short)1);
            retval.Options = ((short)2);
            retval.HResolution = ((short)300);
            retval.VResolution = ((short)300);
            retval.HeaderMargin = (0.5);
            retval.FooterMargin = (0.5);
            retval.Copies = ((short)1);
            return retval;
        }


        /**
         * Returns the HeaderRecord.
         * @return HeaderRecord for the sheet.
         */
        public HeaderRecord Header
        {
            get
            {
                return header;
            }
            set 
            {
                header = value;
            }
        }

        /**
         * Returns the FooterRecord.
         * @return FooterRecord for the sheet.
         */
        public FooterRecord Footer
        {
            get
            {
                return footer;
            }
            set { footer = value; }
        }

        /**
         * Returns the PrintSetupRecord.
         * @return PrintSetupRecord for the sheet.
         */
        public PrintSetupRecord PrintSetup
        {
            get
            {
                return printSetup;
            }
            set 
            {
                printSetup = value;
            }
        }


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
                        return .75;
                    case MarginType.RightMargin:
                        return .75;
                    case MarginType.TopMargin:
                        return 1.0;
                    case MarginType.BottomMargin:
                        return 1.0;
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
            m.Margin= size;
        }

        /**
         * Shifts all the page breaks in the range "count" number of rows/columns
         * @param breaks The page record to be shifted
         * @param start Starting "main" value to shift breaks
         * @param stop Ending "main" value to shift breaks
         * @param count number of units (rows/columns) to shift by
         */
        private static void ShiftBreaks(PageBreakRecord breaks, int start, int stop, int count) {

		IEnumerator iterator = breaks.GetBreaksEnumerator();
		IList shiftedBreak = new ArrayList();
		while(iterator.MoveNext())
		{
			PageBreakRecord.Break breakItem = (PageBreakRecord.Break)iterator.Current;
			int breakLocation = breakItem.main;
			bool inStart = (breakLocation >= start);
			bool inEnd = (breakLocation <= stop);
			if(inStart && inEnd)
				shiftedBreak.Add(breakItem);
		}

		iterator = shiftedBreak.GetEnumerator();
		while (iterator.MoveNext()) {
			PageBreakRecord.Break breakItem = (PageBreakRecord.Break)iterator.Current;
			breaks.RemoveBreak(breakItem.main);
			breaks.AddBreak((short)(breakItem.main+count), breakItem.subFrom, breakItem.subTo);
		}
	}


        /**
         * Sets a page break at the indicated row
         * @param row
         */
        public void SetRowBreak(int row, short fromCol, short toCol)
        {
            this.RowBreaksRecord.AddBreak((short)row, fromCol, toCol);
        }

        /**
         * Removes a page break at the indicated row
         * @param row
         */
        public void RemoveRowBreak(int row)
        {
            if (this.RowBreaksRecord.GetBreaks().Length < 1)
                throw new ArgumentException("Sheet does not define any row breaks");
            this.RowBreaksRecord.RemoveBreak((short)row);
        }

        /**
         * Queries if the specified row has a page break
         * @param row
         * @return true if the specified row has a page break
         */
        public bool IsRowBroken(int row)
        {
            return this.RowBreaksRecord.GetBreak(row) != null;
        }


        /**
         * Queries if the specified column has a page break
         *
         * @return <c>true</c> if the specified column has a page break
         */
        public bool IsColumnBroken(int column)
        {
            return this.ColumnBreaksRecord.GetBreak(column) != null;
        }

        /**
         * Shifts the horizontal page breaks for the indicated count
         * @param startingRow
         * @param endingRow
         * @param count
         */
        public void ShiftRowBreaks(int startingRow, int endingRow, int count)
        {
            ShiftBreaks(this.RowBreaksRecord, startingRow, endingRow, count);
        }

        /**
         * Shifts the vertical page breaks for the indicated count
         * @param startingCol
         * @param endingCol
         * @param count
         */
        public void ShiftColumnBreaks(short startingCol, short endingCol, short count)
        {
            ShiftBreaks(this.ColumnBreaksRecord, startingCol, endingCol, count);
        }

        /**
         * @return all the horizontal page breaks, never <c>null</c>
         */
        public int[] RowBreaks
        {
            get
            {
                return this.RowBreaksRecord.GetBreaks();
            }
        }

        /**
         * @return the number of row page breaks
         */
        public int NumRowBreaks
        {
            get
            {
                return this.RowBreaksRecord.NumBreaks;
            }
        }

        /**
         * @return all the column page breaks, never <c>null</c>
         */
        public int[] ColumnBreaks
        {
            get
            {
                return this.ColumnBreaksRecord.GetBreaks();
            }
        }

        /**
         * @return the number of column page breaks
         */
        public int NumColumnBreaks
        {
            get
            {
                return this.ColumnBreaksRecord.NumBreaks;
            }
        }

        public VCenterRecord VCenter
        {
            get { return _vCenter; }
        }

        public HCenterRecord HCenter
        {
            get { return _hCenter; }
        }
        /// <summary>
        ///  HEADERFOOTER is new in 2007.  Some apps seem to have scattered this record long after
        /// the PageSettingsBlock where it belongs.
        /// </summary>
        /// <param name="rec"></param>
        public void AddLateHeaderFooter(HeaderFooterRecord rec)
        {
            if (_headerFooter != null)
            {
                throw new ArgumentNullException("This page settings block already has a header/footer record");
            }
            if (rec.Sid != UnknownRecord.HEADER_FOOTER_089C)
            {
                throw new RecordFormatException("Unexpected header-footer record sid: 0x" + StringUtil.ToHexString(rec.Sid));
            }
            _headerFooter = rec;
        }


        /// <summary>
        /// This method reads PageSettingsBlock records from the supplied RecordStream until the first non-PageSettingsBlock record is encountered.
        /// As each record is read, it is incorporated into this PageSettingsBlock.
        /// </summary>
        /// <param name="rs"></param> 
        public void AddLateRecords(RecordStream rs)
        {
            while (true)
            {
                if (!ReadARecord(rs))
                {
                    break;
                }
            }
        }
        public void PositionRecords(List<RecordBase> sheetRecords)
        {
            // Take a copy to loop over, so we can update the real one
            //  without concurrency issues
            List<HeaderFooterRecord> hfRecordsToIterate = new List<HeaderFooterRecord>(_sviewHeaderFooters);
            Dictionary<String, HeaderFooterRecord> hfGuidMap = new Dictionary<String, HeaderFooterRecord>();

            foreach (HeaderFooterRecord hf in hfRecordsToIterate)
            {
                string key = HexDump.ToHex(hf.Guid);
                if (hfGuidMap.ContainsKey(key))
                    hfGuidMap[key] = hf;
                else
                    hfGuidMap.Add(HexDump.ToHex(hf.Guid), hf);
            }

            // loop through HeaderFooterRecord records having not-empty GUID and match them with
            // CustomViewSettingsRecordAggregate blocks having UserSViewBegin with the same GUID
            foreach (HeaderFooterRecord hf in hfRecordsToIterate)
            {
                foreach (RecordBase rb in sheetRecords)
                {
                    if (rb is CustomViewSettingsRecordAggregate)
                    {
                        CustomViewSettingsRecordAggregate cv = (CustomViewSettingsRecordAggregate)rb;
                        cv.VisitContainedRecords(new CustomRecordVisitor1(cv,hf,_sviewHeaderFooters,hfGuidMap));
                    }
                }
            }
        }
        private class CustomRecordVisitor1 : RecordVisitor
        {
            CustomViewSettingsRecordAggregate _cv;
            HeaderFooterRecord _hf;
            List<HeaderFooterRecord> _sviewHeaderFooters;
            Dictionary<String, HeaderFooterRecord> _hfGuidMap;
            public CustomRecordVisitor1(CustomViewSettingsRecordAggregate cv, HeaderFooterRecord hf, 
                List<HeaderFooterRecord> sviewHeaderFooter, Dictionary<String, HeaderFooterRecord> hfGuidMap)
            {
                this._cv = cv;
                this._hf = hf;
                this._sviewHeaderFooters = sviewHeaderFooter;
                _hfGuidMap = hfGuidMap;
            }

            #region RecordVisitor Members

            public void VisitRecord(Record r)
            {
                if (r.Sid == UserSViewBegin.sid)
                {
                    String guid = HexDump.ToHex(((UserSViewBegin) r).Guid);
                    HeaderFooterRecord hf = _hfGuidMap[guid];

                    if (hf != null)
                    {
                        {
                            _cv.Append(_hf);
                            _sviewHeaderFooters.Remove(_hf);
                        }
                    }
                }
            }

            #endregion
        }

    }
}
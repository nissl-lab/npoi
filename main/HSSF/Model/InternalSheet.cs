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

namespace NPOI.HSSF.Model
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.Formula;
    using NPOI.SS.Util;
    using NPOI.Util;

    /// <summary>
    /// Low level model implementation of a Sheet (one workbook Contains many sheets)
    /// This file Contains the low level binary records starting at the sheets BOF and
    /// ending with the sheets EOF.  Use HSSFSheet for a high level representation.
    /// 
    /// The structures of the highlevel API use references to this to perform most of their
    /// operations.  Its probably Unwise to use these low level structures directly Unless you
    /// really know what you're doing.  I recommend you Read the Microsoft Excel 97 Developer's
    /// Kit (Microsoft Press) and the documentation at http://sc.openoffice.org/excelfileformat.pdf
    /// before even attempting to use this.
    /// </summary>
    /// <remarks>
    /// @author  Andrew C. Oliver (acoliver at apache dot org)
    /// @author  Glen Stampoultzis (glens at apache.org)
    /// @author  Shawn Laubach (slaubach at apache dot org) Gridlines, Headers, Footers, PrintSetup, and Setting Default Column Styles
    /// @author Jason Height (jheight at chariot dot net dot au) Clone support. DBCell and Index Record writing support
    /// @author  Brian Sanders (kestrel at burdell dot org) Active Cell support
    /// @author  Jean-Pierre Paris (jean-pierre.paris at m4x dot org) (Just a little)
    /// </remarks>
    [Serializable]
    public class InternalSheet
    {

        //private static POILogger log = POILogFactory.GetLogger(typeof(Sheet));

        int preoffset = 0;            // offset of the sheet in a new file
        protected int dimsloc = -1;  // TODO - Is it legal for dims record to be missing?
        [NonSerialized]
        protected DimensionsRecord dims;
        [NonSerialized]
        protected DefaultColWidthRecord defaultcolwidth = new DefaultColWidthRecord();
        [NonSerialized]
        protected DefaultRowHeightRecord defaultrowheight = new DefaultRowHeightRecord();
        [NonSerialized]
        protected GridsetRecord gridset = null;
        [NonSerialized]
        protected PrintSetupRecord printSetup = null;
        [NonSerialized]
        protected HeaderRecord header = null;
        [NonSerialized]
        protected FooterRecord footer = null;
        [NonSerialized]
        protected PrintGridlinesRecord printGridlines = null;
        [NonSerialized]
        protected WindowTwoRecord windowTwo = null;
        [NonSerialized]
        protected MergeCellsRecord merged = null;
        /** java object always present, but if empty no BIFF records are written */
        [NonSerialized]
        private MergedCellsTable _mergedCellsTable;
        [NonSerialized]
        protected RowRecordsAggregate _rowsAggregate;
        [NonSerialized]
        private PageSettingsBlock _psBlock;
        protected IMargin[] margins = null;
        //protected IList mergedRecords = new ArrayList();

        [NonSerialized]
        protected SelectionRecord selection = null;
        //protected ColumnInfoRecordsAggregate columns = null;
        //protected ValueRecordsAggregate cells = null;
        /*package*/
        [NonSerialized]
        internal ColumnInfoRecordsAggregate _columnInfos;
        /** the DimensionsRecord is always present */
        [NonSerialized]
        private DimensionsRecord _dimensions;
        [NonSerialized]
        private DataValidityTable _dataValidityTable = null;
        //private IEnumerator valueRecEnumerator = null;
        private IEnumerator rowRecEnumerator = null;
        protected int eofLoc = 0;
        [NonSerialized]
        private GutsRecord _gutsRecord;
        [NonSerialized]
        protected PageBreakRecord rowBreaks = null;
        [NonSerialized]
        protected PageBreakRecord colBreaks = null;
        [NonSerialized]
        protected ConditionalFormattingTable condFormatting;
        [NonSerialized]
        protected SheetExtRecord sheetext;
        protected List<RecordBase> records = null;

        /** Add an UncalcedRecord if not true indicating formulas have not been calculated */
        protected bool _isUncalced = false;

        //public const byte PANE_LOWER_RIGHT = (byte)0;
        //public const byte PANE_UPPER_RIGHT = (byte)1;
        //public const byte PANE_LOWER_LEFT = (byte)2;
        //public const byte PANE_UPPER_LEFT = (byte)3;


        /// <summary>
        /// Clones the low level records of this sheet and returns the new sheet instance.
        /// This method is implemented by Adding methods for deep cloning to all records that
        /// can be Added to a sheet. The Record object does not implement Cloneable.
        /// When Adding a new record, implement a public Clone method if and only if the record
        /// belongs to a sheet.
        /// </summary>
        /// <returns></returns>
        public InternalSheet CloneSheet()
        {
            List<RecordBase> clonedRecords = new List<RecordBase>(this.records.Count);
            for (int i = 0; i < this.records.Count; i++)
            {
                RecordBase rb = (RecordBase)this.records[i];
                if (rb is RecordAggregate)
                {
                    ((RecordAggregate)rb).VisitContainedRecords(new RecordCloner(clonedRecords));
                    continue;
                }

                if (rb is EscherAggregate)
                {
                    /*
                     * this record will be removed after reading actual data from EscherAggregate
                     */
                    rb = new DrawingRecord();
                }
                Record rec = (Record)((Record)rb).Clone();
                clonedRecords.Add(rec);

                
            }
            return CreateSheet(new RecordStream(clonedRecords, 0));
        }
        /// <summary>
        /// get the NEXT value record (from LOC).  The first record that is a value record
        /// (starting at LOC) will be returned.
        /// This method is "loc" sensitive.  Meaning you need to set LOC to where you
        /// want it to start searching.  If you don't know do this: setLoc(getDimsLoc).
        /// When adding several rows you can just start at the last one by leaving loc
        /// at what this sets it to.  For this method, set loc to dimsloc to start with,
        /// subsequent calls will return values in (physical) sequence or NULL when you get to the end.
        /// </summary>
        /// <returns>the next value record or NULL if there are no more</returns>
        // <see cref="SetLoc(int)"/>
        public CellValueRecordInterface[] GetValueRecords()
        {
            return _rowsAggregate.GetValueRecords();
        }

        public WindowTwoRecord WindowTwo
        {
            get { return windowTwo; }
        }
        /// <summary>
        /// Creates the sheet.
        /// </summary>
        /// <param name="rs">The stream.</param>
        /// <returns></returns>
        public static InternalSheet CreateSheet(RecordStream rs)
        {
            return new InternalSheet(rs);
        }

        private class RecordCloner : RecordVisitor
        {
            private IList _destList;

            public RecordCloner(IList destList)
            {
                _destList = destList;
            }
            public void VisitRecord(Record r)
            {
                _destList.Add(r.Clone());
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="InternalSheet"/> class.
        /// </summary>
        /// <param name="rs">The stream.</param>
        private InternalSheet(RecordStream rs)
        {
            _mergedCellsTable = new MergedCellsTable();
            RowRecordsAggregate rra = null;

            records = new List<RecordBase>(128);
            // TODO - take chart streams off into separate java objects
            //int bofEofNestingLevel = 0;  // nesting level can only get to 2 (when charts are present)
            int dimsloc = -1;

            if (rs.PeekNextSid() != BOFRecord.sid)
            {
                throw new RuntimeException("BOF record expected");
            }
            BOFRecord bof = (BOFRecord)rs.GetNext();
            if (bof.Type == BOFRecordType.Worksheet)
            {
                // Good, well supported
            }
            else if (bof.Type == BOFRecordType.Chart ||
                     bof.Type == BOFRecordType.Excel4Macro)
            {
                // These aren't really typical sheets... Let it go though,
                //  we can handle them roughly well enough as a "normal" one
            }
            else
            {
                // Not a supported type
                // Skip onto the EOF, then complain
                while (rs.HasNext())
                {
                    Record rec = rs.GetNext();
                    if (rec is EOFRecord)
                    {
                        break;
                    }
                }
                throw new UnsupportedBOFType(bof.Type);
            }
        
            records.Add(bof);
            while (rs.HasNext())
            {
                int recSid = rs.PeekNextSid();

                if (recSid == CFHeaderRecord.sid)
                {
                    condFormatting = new ConditionalFormattingTable(rs);
                    records.Add(condFormatting);
                    continue;
                }

                if (recSid == ColumnInfoRecord.sid)
                {
                    _columnInfos = new ColumnInfoRecordsAggregate(rs);
                    records.Add(_columnInfos);
                    continue;
                }
                if (recSid == DVALRecord.sid)
                {
                    _dataValidityTable = new DataValidityTable(rs);
                    records.Add(_dataValidityTable);
                    continue;
                }

                if (RecordOrderer.IsRowBlockRecord(recSid))
                {
                    //only Add the aggregate once
                    if (rra != null)
                    {
                        throw new InvalidOperationException("row/cell records found in the wrong place");
                    }
                    RowBlocksReader rbr = new RowBlocksReader(rs);
                    _mergedCellsTable.AddRecords(rbr.LooseMergedCells);
                    rra = new RowRecordsAggregate(rbr.PlainRecordStream, rbr.SharedFormulaManager);
                    records.Add(rra); //only Add the aggregate once
                    continue;
                }

                if (CustomViewSettingsRecordAggregate.IsBeginRecord(recSid))
                {
                    // This happens three times in test sample file "29982.xls"
                    // Also several times in bugzilla samples 46840-23373 and 46840-23374
                    records.Add(new CustomViewSettingsRecordAggregate(rs));
                    continue;                    
                }

                if (PageSettingsBlock.IsComponentRecord(recSid))
                {
                    if (_psBlock == null)
                    {
                        // first PSB record encountered - read all of them:
                        _psBlock = new PageSettingsBlock(rs);
                        records.Add(_psBlock);
                    }
                    else
                    {
                        // one or more PSB records found after some intervening non-PSB records
                        _psBlock.AddLateRecords(rs);
                    }
                    // YK: in some cases records can be moved to the preceding
                    // CustomViewSettingsRecordAggregate blocks
                    _psBlock.PositionRecords(records);
                    continue;
                }
                if (WorksheetProtectionBlock.IsComponentRecord(recSid))
                {
                    _protectionBlock.AddRecords(rs);
                    continue;
                }
                if (recSid == MergeCellsRecord.sid)
                {
                    // when the MergedCellsTable is found in the right place, we expect those records to be contiguous
                    _mergedCellsTable.Read(rs);
                    continue;
                }
                
                if (recSid == BOFRecord.sid)
                {
                    ChartSubstreamRecordAggregate chartAgg = new ChartSubstreamRecordAggregate(rs);
                    //ChartSheetAggregate chartAgg = new ChartSheetAggregate(rs, null);
                    //if (false)
                    //{
                    // TODO - would like to keep the chart aggregate packed, but one unit test needs attention
                    //    records.Add(chartAgg);
                    //}
                    //else
                    //{
                    SpillAggregate(chartAgg, records);
                    //}
                    continue;
                }
                Record rec = rs.GetNext();
                if (recSid == IndexRecord.sid)
                {
                    // ignore INDEX record because it is only needed by Excel,
                    // and POI always re-calculates its contents
                    continue;
                }

                
                if (recSid == UncalcedRecord.sid)
                {
                    // don't Add UncalcedRecord to the list
                    _isUncalced = true; // this flag is enough
                    continue;
                }

                if (recSid == FeatRecord.sid ||
                    recSid == FeatHdrRecord.sid)
                {
                    records.Add(rec);
                    continue;
                }
                if (recSid == EOFRecord.sid)
                {
                    records.Add(rec);
                    break;
                }
                if (recSid == DimensionsRecord.sid)
                {
                    // Make a columns aggregate if one hasn't Ready been created.
                    if (_columnInfos == null)
                    {
                        _columnInfos = new ColumnInfoRecordsAggregate();
                        records.Add(_columnInfos);
                    }

                    _dimensions = (DimensionsRecord)rec;
                    dimsloc = records.Count;
                }
                else if (recSid == DefaultColWidthRecord.sid)
                {
                    defaultcolwidth = (DefaultColWidthRecord)rec;
                }
                else if (recSid == DefaultRowHeightRecord.sid)
                {
                    defaultrowheight = (DefaultRowHeightRecord)rec;
                }
                else if (recSid == PrintGridlinesRecord.sid)
                {
                    printGridlines = (PrintGridlinesRecord)rec;
                }
                else if (recSid == GridsetRecord.sid)
                {
                    gridset = (GridsetRecord)rec;
                }
                else if (recSid == SelectionRecord.sid)
                {
                    selection = (SelectionRecord)rec;
                }
                else if (recSid == WindowTwoRecord.sid)
                {
                    windowTwo = (WindowTwoRecord)rec;
                }
                else if (recSid == SheetExtRecord.sid)
                {
                    sheetext = (SheetExtRecord)rec;
                }
                else if (recSid == GutsRecord.sid)
                {
                    _gutsRecord = (GutsRecord)rec;
                }

                records.Add(rec);
            }
            if (windowTwo == null)
            {
                throw new InvalidOperationException("WINDOW2 was not found");
            }
            if (_dimensions == null)
            {
                // Excel seems to always write the DIMENSION record, but tolerates when it is not present
                // in all cases Excel (2007) adds the missing DIMENSION record
                if (rra == null)
                {
                    // bug 46206 alludes to files which skip the DIMENSION record
                    // when there are no row/cell records.
                    // Not clear which application wrote these files.
                    rra = new RowRecordsAggregate();
                }
                else
                {
                    //log.log(POILogger.WARN, "DIMENSION record not found even though row/cells present");
                    // Not sure if any tools write files like this, but Excel reads them OK
                }
                dimsloc = FindFirstRecordLocBySid(WindowTwoRecord.sid);
                _dimensions = rra.CreateDimensions();
                records.Insert(dimsloc, _dimensions);                
            }
            if (rra == null)
            {
                rra = new RowRecordsAggregate();
                records.Insert(dimsloc + 1, rra);
            }
            _rowsAggregate = rra;
            // put merged cells table in the right place (regardless of where the first MergedCellsRecord was found */
            RecordOrderer.AddNewSheetRecord(records, _mergedCellsTable);
            RecordOrderer.AddNewSheetRecord(records, _protectionBlock);
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "sheet createSheet (existing file) exited");

        }
        private class RecordVisitor1:RecordVisitor
        {
            List<RecordBase> _records;
            public RecordVisitor1(List<RecordBase> recs)
            {
                _records=recs;
            }
            public void VisitRecord(Record r) {
                _records.Add(r);
            }
        }
        private static void SpillAggregate(RecordAggregate ra, List<RecordBase> recs) {
            ra.VisitContainedRecords(new RecordVisitor1(recs));
        }
        /// <summary>
        /// Creates a sheet with all the usual records minus values and the "index"
        /// record (not required).  Sets the location pointer to where the first value
        /// records should go.  Use this to Create a sheet from "scratch".
        /// </summary>
        /// <returns>Sheet object with all values Set to defaults</returns>
        public static InternalSheet CreateSheet()
        {
            return new InternalSheet();
        }

        private InternalSheet()
        {
            _mergedCellsTable = new MergedCellsTable();
            records = new List<RecordBase>(32);

            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "Sheet Createsheet from scratch called");

            records.Add(CreateBOF());

            records.Add(CreateCalcMode());
            records.Add(CreateCalcCount());
            records.Add(CreateRefMode());
            records.Add(CreateIteration());
            records.Add(CreateDelta());
            records.Add(CreateSaveRecalc());
            records.Add(CreatePrintHeaders());
            printGridlines = CreatePrintGridlines();
            records.Add(printGridlines);
            gridset = CreateGridset();
            records.Add(gridset);
            _gutsRecord = CreateGuts();
            records.Add(_gutsRecord);
            defaultrowheight = CreateDefaultRowHeight();
            records.Add(defaultrowheight);
            records.Add(CreateWSBool());

            // 'Page Settings Block'
            _psBlock = new PageSettingsBlock();
            records.Add(_psBlock);

            // 'Worksheet Protection Block' (after 'Page Settings Block' and before DEFCOLWIDTH)
            records.Add(_protectionBlock);

            defaultcolwidth = CreateDefaultColWidth();
            records.Add(defaultcolwidth);
            ColumnInfoRecordsAggregate columns = new ColumnInfoRecordsAggregate();
            records.Add(columns);
            _columnInfos = columns;
            _dimensions = CreateDimensions();
            records.Add(_dimensions);
            _rowsAggregate = new RowRecordsAggregate();
            records.Add(_rowsAggregate);
            // 'Sheet View Settings'
            records.Add(windowTwo = CreateWindowTwo());
            selection = CreateSelection();
            records.Add(selection);

            records.Add(_mergedCellsTable); // MCT comes after 'Sheet View Settings'
            sheetext = new SheetExtRecord();
            records.Add(sheetext);
            records.Add(EOFRecord.instance);

            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "Sheet Createsheet from scratch exit");
        }

        /// <summary>
        /// Adds the merged region.
        /// </summary>
        /// <param name="rowFrom">the row index From </param>
        /// <param name="colFrom">The column index From.</param>
        /// <param name="rowTo">The row index To</param>
        /// <param name="colTo">The column To.</param>
        /// <returns></returns>
        public int AddMergedRegion(int rowFrom, int colFrom, int rowTo, int colTo)
        {
            // Validate input
            if (rowTo < rowFrom)
            {
                throw new ArgumentException("The 'to' row (" + rowTo
                        + ") must not be less than the 'from' row (" + rowFrom + ")");
            }
            if (colTo < colFrom)
            {
                throw new ArgumentException("The 'to' col (" + colTo
                        + ") must not be less than the 'from' col (" + colFrom + ")");
            }

            MergedCellsTable mrt = MergedRecords;
            mrt.AddArea(rowFrom, colFrom, rowTo, colTo);
            return mrt.NumberOfMergedRegions - 1;
        }

        /// <summary>
        /// Removes the merged region.
        /// </summary>
        /// <param name="index">The index.</param>
        public void RemoveMergedRegion(int index)
        {
            //safety checks
            MergedCellsTable mrt = MergedRecords;
            if (index >= mrt.NumberOfMergedRegions)
            {
                return;
            }
            mrt.Remove(index);
        }
        /// <summary>
        /// Gets the column infos.
        /// </summary>
        /// <value>The column infos.</value>
        public ColumnInfoRecordsAggregate ColumnInfos
        {
            get { return _columnInfos; }
        }
        internal MergedCellsTable MergedRecords
        {
            get
            {
                // always present
                return _mergedCellsTable;
            }
        }


        /// <summary>
        /// Gets the merged region at.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public CellRangeAddress GetMergedRegionAt(int index)
        {
            //safety checks
            MergedCellsTable mrt = MergedRecords;
            if (index >= mrt.NumberOfMergedRegions)
            {
                return null;
            }
            return mrt.Get(index);
        }

        /// <summary>
        /// Gets the number of merged regions.
        /// </summary>
        /// <value>The number merged regions.</value>
        public int NumMergedRegions
        {
            get
            {
                return MergedRecords.NumberOfMergedRegions;
            }
        }
        ///// <summary>
        ///// Find correct position to Add new CF record
        ///// </summary>
        ///// <returns></returns>
        //private int FindConditionalFormattingPosition()
        //{
        //    // This is default.
        //    // If the algorithm does not Find the right position,
        //    // this one will be used (this is a position before EOF record)
        //    int index = records.Count - 2;

        //    for (int i = index; i >= 0; i--)
        //    {
        //        Record rec = (Record)records[i];
        //        short sid = rec.Sid;

        //        // CFRecordsAggregate records already exist, just Add to the end
        //        if (rec is CFRecordsAggregate) { return i + 1; }

        //        if (sid == (short)0x00ef) { return i + 1; }// PHONETICPR
        //        if (sid == (short)0x015f) { return i + 1; }// LABELRANGES
        //        if (sid == MergeCellsRecord.sid) { return i + 1; }
        //        if (sid == (short)0x0099) { return i + 1; }// STANDARDWIDTH
        //        if (sid == SelectionRecord.sid) { return i + 1; }
        //        if (sid == PaneRecord.sid) { return i + 1; }
        //        if (sid == SCLRecord.sid) { return i + 1; }
        //        if (sid == WindowTwoRecord.sid) { return i + 1; }
        //    }

        //    return index;
        //}


        /// <summary>
        /// Gets the number of conditional formattings.
        /// </summary>
        /// <value>The number of conditional formattings.</value>
        public int NumConditionalFormattings
        {
            get
            {
                return condFormatting.Count;
            }
        }

        /// <summary>
        /// Per an earlier reported bug in working with Andy Khan's excel Read library.  This
        /// Sets the values in the sheet's DimensionsRecord object to be correct.  Excel doesn't
        /// really care, but we want to play nice with other libraries.
        /// </summary>
        /// <param name="firstrow">The first row.</param>
        /// <param name="firstcol">The first column.</param>
        /// <param name="lastrow">The last row.</param>
        /// <param name="lastcol">The last column.</param>
        public void SetDimensions(int firstrow, short firstcol, int lastrow,
                                  short lastcol)
        {
            //if (log.Check(POILogger.DEBUG))
            //{
            //    log.Log(POILogger.DEBUG, "Sheet.SetDimensions");
            //    log.Log(POILogger.DEBUG,
            //            (new StringBuilder("firstrow")).Append(firstrow)
            //                .Append("firstcol").Append(firstcol).Append("lastrow")
            //                .Append(lastrow).Append("lastcol").Append(lastcol)
            //                .ToString());
            //}
            dims.FirstCol = firstcol;
            dims.FirstRow = firstrow;
            dims.LastCol = lastcol;
            dims.LastRow = lastrow;
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "Sheet.SetDimensions exiting");
        }

        /// <summary>
        /// Gets or Sets the preoffset when using DBCELL records (currently Unused) - this Is
        /// the position of this sheet within the whole file.
        /// </summary>
        /// <value>the offset of the sheet's BOF within the file.</value>
        public int PreOffset
        {
            get
            {
                return preoffset;
            }
            set
            {
                this.preoffset = value;
            }
        }
        /// <summary>
        /// Create a row record.  (does not Add it to the records contained in this sheet)
        /// </summary>
        /// <param name="row">row number</param>
        /// <returns>RowRecord Created for the passed in row number</returns>
        public RowRecord CreateRow(int row)
        {
            return RowRecordsAggregate.CreateRow(row);
        }

        /// <summary>
        /// Create a LABELSST Record (does not Add it to the records contained in this sheet)
        /// </summary>
        /// <param name="row">the row the LabelSST Is a member of</param>
        /// <param name="col">the column the LabelSST defines</param>
        /// <param name="index">the index of the string within the SST (use workbook AddSSTString method)</param>
        /// <returns>LabelSSTRecord newly Created containing your SST Index, row,col.</returns>
        public LabelSSTRecord CreateLabelSST(int row, short col, int index)
        {
            //    log.LogFormatted(POILogger.DEBUG, "Create labelsst row,col,index %,%,%",
            //                     new int[]
            //{
            //    row, col, index
            //});
            LabelSSTRecord rec = new LabelSSTRecord();

            rec.Row = (row);
            rec.Column = (col);
            rec.SSTIndex = (index);
            rec.XFIndex = ((short)0x0f);
            return rec;
        }

        /// <summary>
        /// Create a NUMBER Record (does not Add it to the records contained in this sheet)
        /// </summary>
        /// <param name="row">the row the NumberRecord is a member of</param>
        /// <param name="col">the column the NumberRecord defines</param>
        /// <param name="value">value for the number record</param>
        /// <returns>NumberRecord for that row, col containing that value as Added to the sheet</returns>
        public NumberRecord CreateNumber(int row, short col, double value)
        {
            //log.LogFormatted(POILogger.DEBUG, "Create number row,col,value %,%,%",
            //                 new double[]
            //{
            //    row, col, value
            //});
            NumberRecord rec = new NumberRecord();

            rec.Row = row;
            rec.Column = col;
            rec.Value = value;
            rec.XFIndex = (short)0x0f;
            return rec;
        }

        /// <summary>
        /// Create a BLANK record (does not Add it to the records contained in this sheet)
        /// </summary>
        /// <param name="row">the row the BlankRecord is a member of</param>
        /// <param name="col">the column the BlankRecord is a member of</param>
        /// <returns></returns>
        public BlankRecord CreateBlank(int row, short col)
        {
            //log.LogFormatted(POILogger.DEBUG, "Create blank row,col %,%", new int[]
            //{
            //    row, col
            //});
            BlankRecord rec = new BlankRecord();

            rec.Row = row;
            rec.Column = col;
            rec.XFIndex = (short)0x0f;
            return rec;
        }

        /// <summary>
        /// Adds a value record to the sheet's contained binary records
        /// (i.e. LabelSSTRecord or NumberRecord).
        /// This method is "loc" sensitive.  Meaning you need to Set LOC to where you
        /// want it to start searching.  If you don't know do this: SetLoc(GetDimsLoc).
        /// When Adding several rows you can just start at the last one by leaving loc
        /// at what this Sets it to.
        /// </summary>
        /// <param name="row">the row to Add the cell value to</param>
        /// <param name="col">the cell value record itself.</param>
        public void AddValueRecord(int row, CellValueRecordInterface col)
        {
            //if (log.Check(POILogger.DEBUG))
            //{
            //    log.Log(POILogger.DEBUG, "Add value record  row" + row);
            //}
            DimensionsRecord d = _dimensions;

            if (col.Column >= d.LastCol)
            {
                d.LastCol = ((short)(col.Column + 1));
            }
            if (col.Column < d.FirstCol)
            {
                d.FirstCol = (col.Column);
            }
            _rowsAggregate.InsertCell(col);
        }

        /// <summary>
        /// Remove a value record from the records array.
        /// This method is not loc sensitive, it Resets loc to = dimsloc so no worries.
        /// </summary>
        /// <param name="row">the row of the value record you wish to Remove</param>
        /// <param name="col">a record supporting the CellValueRecordInterface.</param>
        public void RemoveValueRecord(int row, CellValueRecordInterface col)
        {
            //log.LogFormatted(POILogger.DEBUG, "Remove value record row,dimsloc %,%",
            //                 new int[] { row, dimsloc });
            _rowsAggregate.RemoveCell(col);
        }

        /// <summary>
        /// Replace a value record from the records array.
        /// This method is not loc sensitive, it Resets loc to = dimsloc so no worries.
        /// </summary>
        /// <param name="newval">a record supporting the CellValueRecordInterface.  this will Replace
        /// the cell value with the same row and column.  If there Isn't one, one will
        /// be Added.</param>
        public void ReplaceValueRecord(CellValueRecordInterface newval)
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "ReplaceValueRecord ");
            //The ValueRecordsAggregate use a tree map Underneath.
            //The tree Map uses the CellValueRecordInterface as both the
            //key and the value, if we dont do a Remove, then
            //the previous instance of the key is retained, effectively using
            //double the memory
            _rowsAggregate.RemoveCell(newval);
            _rowsAggregate.InsertCell(newval);
        }

        /// <summary>
        /// Adds a row record to the sheet
        /// This method is "loc" sensitive.  Meaning you need to Set LOC to where you
        /// want it to start searching.  If you don't know do this: SetLoc(GetDimsLoc).
        /// When Adding several rows you can just start at the last one by leaving loc
        /// at what this Sets it to.
        /// </summary>
        /// <param name="row">the row record to be Added</param>
        public void AddRow(RowRecord row)
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "AddRow ");
            DimensionsRecord d = _dimensions;

            if (row.RowNumber >= d.LastRow)
            {
                d.LastRow = (row.RowNumber + 1);
            }
            if (row.RowNumber < d.FirstRow)
            {
                d.FirstRow = (row.RowNumber);
            }
            //IndexRecord index = null;
            //If the row exists Remove it, so that any cells attached to the row are Removed
            RowRecord existingRow = _rowsAggregate.GetRow(row.RowNumber);
            if (existingRow != null)
            {
                _rowsAggregate.RemoveRow(existingRow);
            }

            _rowsAggregate.InsertRow(row);

            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "exit AddRow");
        }

        /// <summary>
        /// Removes a row record
        /// This method is not loc sensitive, it Resets loc to = dimsloc so no worries.
        /// </summary>
        /// <param name="row">the row record to Remove</param>
        public void RemoveRow(RowRecord row)
        {
            _rowsAggregate.RemoveRow(row);
        }


        /// <summary>
        /// Get the NEXT RowRecord (from LOC).  The first record that is a Row record
        /// (starting at LOC) will be returned.
        /// This method is "loc" sensitive.  Meaning you need to Set LOC to where you
        /// want it to start searching.  If you don't know do this: SetLoc(GetDimsLoc).
        /// When Adding several rows you can just start at the last one by leaving loc
        /// at what this Sets it to.  For this method, Set loc to dimsloc to start with.
        /// subsequent calls will return rows in (physical) sequence or NULL when you Get to the end.
        /// </summary>
        /// <value>RowRecord representing the next row record or NULL if there are no more</value>
        public RowRecord NextRow
        {
            get
            {
                if (rowRecEnumerator == null)
                {
                    rowRecEnumerator = _rowsAggregate.GetEnumerator();
                }
                if (!rowRecEnumerator.MoveNext())
                {
                    return null;
                }
                return (RowRecord)rowRecEnumerator.Current;
            }
        }

        /// <summary>
        /// Get the NEXT (from LOC) RowRecord where rownumber matches the given rownum.
        /// The first record that is a Row record (starting at LOC) that has the
        /// same rownum as the given rownum will be returned.
        /// This method is "loc" sensitive.  Meaning you need to Set LOC to where you
        /// want it to start searching.  If you don't know do this: SetLoc(GetDimsLoc).
        /// When Adding several rows you can just start at the last one by leaving loc
        /// at what this Sets it to.  For this method, Set loc to dimsloc to start with.
        /// subsequent calls will return rows in (physical) sequence or NULL when you Get to the end.
        /// </summary>
        /// <param name="rownum">which row to return (careful with LOC)</param>
        /// <returns>RowRecord representing the next row record or NULL if there are no more</returns>
        public RowRecord GetRow(int rownum)
        {
            return _rowsAggregate.GetRow(rownum);
        }
        /// <summary>
        /// Gets the page settings.
        /// </summary>
        /// <returns></returns>
        public PageSettingsBlock PageSettings
        {
            get
            {
                if (_psBlock == null)
                {
                    _psBlock = new PageSettingsBlock();
                    RecordOrderer.AddNewSheetRecord(records, _psBlock);
                }
                return _psBlock;
            }
        }
        /// <summary>
        /// Creates the BOF record
        /// </summary>
        /// <returns>record containing a BOFRecord</returns>
        public static Record CreateBOF()
        {
            BOFRecord retval = new BOFRecord();

            retval.Version = ((short)0x600);
            retval.Type = BOFRecordType.Worksheet;

            retval.Build = ((short)0x0dbb);
            retval.BuildYear = ((short)1996);
            retval.HistoryBitMask = (0xc1);
            retval.RequiredVersion = (0x6);
            return retval;
        }

        /// <summary>
        /// Creates the Index record  - not currently used
        /// </summary>
        /// <returns>record containing a IndexRecord</returns>
        protected Record CreateIndex()
        {
            IndexRecord retval = new IndexRecord();

            retval.FirstRow = (0);   // must be Set explicitly
            retval.LastRowAdd1 = (0);
            return retval;
        }

        /// <summary>
        /// Creates the CalcMode record and Sets it to 1 (automatic formula caculation)
        /// </summary>
        /// <returns>record containing a CalcModeRecord</returns>
        protected Record CreateCalcMode()
        {
            CalcModeRecord retval = new CalcModeRecord();

            retval.SetCalcMode((short)1);
            return retval;
        }

        /// <summary>
        /// Creates the CalcCount record and Sets it to 0x64 (default number of iterations)
        /// </summary>
        /// <returns>record containing a CalcCountRecord</returns>
        protected Record CreateCalcCount()
        {
            CalcCountRecord retval = new CalcCountRecord();

            retval.Iterations = ((short)0x64);   // default 64 iterations
            return retval;
        }

        /// <summary>
        /// Creates the RefMode record and Sets it to A1 Mode (default reference mode)
        /// </summary>
        /// <returns>record containing a RefModeRecord</returns>
        protected Record CreateRefMode()
        {
            RefModeRecord retval = new RefModeRecord();

            retval.Mode = RefModeRecord.USE_A1_MODE;
            return retval;
        }

        /// <summary>
        /// Creates the Iteration record and Sets it to false (don't iteratively calculate formulas)
        /// </summary>
        /// <returns>record containing a IterationRecord</returns>
        protected Record CreateIteration()
        {
            return new IterationRecord(false);
        }

        /// <summary>
        /// Creates the Delta record and Sets it to 0.0010 (default accuracy)
        /// </summary>
        /// <returns>record containing a DeltaRecord</returns>
        protected Record CreateDelta()
        {
            return new DeltaRecord(DeltaRecord.DEFAULT_VALUE);            
        }

        /// <summary>
        /// Creates the SaveRecalc record and Sets it to true (recalculate before saving)
        /// </summary>
        /// <returns>record containing a SaveRecalcRecord</returns>
        protected Record CreateSaveRecalc()
        {
            SaveRecalcRecord retval = new SaveRecalcRecord();

            retval.Recalc = (true);
            return retval;
        }

        /// <summary>
        /// Creates the PrintHeaders record and Sets it to false (we don't Create headers yet so why print them)
        /// </summary>
        /// <returns>record containing a PrintHeadersRecord</returns>
        protected Record CreatePrintHeaders()
        {
            PrintHeadersRecord retval = new PrintHeadersRecord();

            retval.PrintHeaders = false;
            return retval;
        }

        /// <summary>
        /// Creates the PrintGridlines record and Sets it to false (that makes for ugly sheets).  As far as I can
        /// tell this does the same thing as the GridsetRecord
        /// </summary>
        /// <returns>record containing a PrintGridlinesRecord</returns>
        protected PrintGridlinesRecord CreatePrintGridlines()
        {
            PrintGridlinesRecord retval = new PrintGridlinesRecord();

            retval.PrintGridlines = false;
            return retval;
        }

        /// <summary>
        /// Creates the GridSet record and Sets it to true (user has mucked with the gridlines)
        /// </summary>
        /// <returns>record containing a GridsetRecord</returns>
        protected GridsetRecord CreateGridset()
        {
            GridsetRecord retval = new GridsetRecord();

            retval.Gridset = (true);
            return retval;
        }

        /// <summary>
        /// Creates the Guts record and Sets leftrow/topcol guttter and rowlevelmax/collevelmax to 0
        /// </summary>
        /// <returns>record containing a GutsRecordRecord</returns>
        protected GutsRecord CreateGuts()
        {
            GutsRecord retval = new GutsRecord();

            retval.LeftRowGutter = ((short)0);
            retval.TopColGutter = ((short)0);
            retval.RowLevelMax = ((short)0);
            retval.ColLevelMax = ((short)0);
            return retval;
        }
        /// <summary>
        /// Creates the DefaultRowHeight Record and Sets its options to 0 and rowheight to 0xff
        /// </summary>
        /// <see cref="NPOI.HSSF.Record.DefaultRowHeightRecord"/>
        /// <see cref="NPOI.HSSF.Record.Record"/>
        /// <returns>record containing a DefaultRowHeightRecord</returns>
        protected DefaultRowHeightRecord CreateDefaultRowHeight()
        {
            DefaultRowHeightRecord retval = new DefaultRowHeightRecord();

            retval.OptionFlags = ((short)0);
            retval.RowHeight = ((short)0xff);
            return retval;
        }

        /**
         * Creates the WSBoolRecord and Sets its values to defaults
         * @see org.apache.poi.hssf.record.WSBoolRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a WSBoolRecord
         */

        protected Record CreateWSBool()
        {
            WSBoolRecord retval = new WSBoolRecord();

            retval.WSBool1 = ((byte)0x4);
            retval.WSBool2 = ((byte)0x1);
            return retval;
        }

        /**
         * Creates the HCenter Record and Sets it to false (don't horizontally center)
         * @see org.apache.poi.hssf.record.HCenterRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a HCenterRecord
         */

        protected Record CreateHCenter()
        {
            HCenterRecord retval = new HCenterRecord();

            retval.HCenter = (false);
            return retval;
        }

        /**
         * Creates the VCenter Record and Sets it to false (don't horizontally center)
         * @see org.apache.poi.hssf.record.VCenterRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a VCenterRecord
         */

        protected Record CreateVCenter()
        {
            VCenterRecord retval = new VCenterRecord();

            retval.VCenter = (false);
            return retval;
        }

        /**
         * Creates the PrintSetup Record and Sets it to defaults and marks it invalid
         * @see org.apache.poi.hssf.record.PrintSetupRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a PrintSetupRecord
         */

        protected Record CreatePrintSetup()
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
            retval.Copies = ((short)0);
            return retval;
        }

        /**
         * Creates the DefaultColWidth Record and Sets it to 8
         * @see org.apache.poi.hssf.record.DefaultColWidthRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a DefaultColWidthRecord
         */

        protected DefaultColWidthRecord CreateDefaultColWidth()
        {
            DefaultColWidthRecord retval = new DefaultColWidthRecord();

            retval.ColWidth = (short)DefaultColWidthRecord.DEFAULT_COLUMN_WIDTH;
            return retval;
        }

        /**
         * Get the default column width for the sheet (if the columns do not define their own width)
         * @return default column width
         */

        public int DefaultColumnWidth
        {
            get { return defaultcolwidth.ColWidth; }
            set { defaultcolwidth.ColWidth = (short)value; }
        }

        /**
         * Get the default row height for the sheet (if the rows do not define their own height)
         * @return  default row height
         */

        public short DefaultRowHeight
        {
            get { return defaultrowheight.RowHeight; }
            set 
            { 
                defaultrowheight.RowHeight = (value);
                // set the bit that specifies that the default settings for the row height have been changed.
                defaultrowheight.OptionFlags = (short)1;
            }
        }

        /**
         * Get the width of a given column in Units of 1/256th of a Char width
         * @param column index
         * @see org.apache.poi.hssf.record.DefaultColWidthRecord
         * @see org.apache.poi.hssf.record.ColumnInfoRecord
         * @see #SetColumnWidth(short,short)
         * @return column width in Units of 1/256th of a Char width
         */

        public int GetColumnWidth(int columnIndex)
        {
            ColumnInfoRecord ci = _columnInfos.FindColumnInfo(columnIndex);
            if (ci != null)
            {
                return ci.ColumnWidth;
            }
            //default column width is measured in characters
            //multiply
            return (256 * defaultcolwidth.ColWidth);
        }

        /**
         * Get the index to the ExtendedFormatRecord "associated" with
         * the column at specified 0-based index. (In this case, an
         * ExtendedFormatRecord index is actually associated with a
         * ColumnInfoRecord which spans 1 or more columns)
         * <br/>
         * Returns the index to the default ExtendedFormatRecord (0xF)
         * if no ColumnInfoRecord exists that includes the column
         * index specified.
         * @param column
         * @return index of ExtendedFormatRecord associated with
         * ColumnInfoRecord that includes the column index or the
         * index of the default ExtendedFormatRecord (0xF)
         */
        public short GetXFIndexForColAt(short columnIndex)
        {
            ColumnInfoRecord ci = _columnInfos.FindColumnInfo(columnIndex);
            if (ci != null)
            {
                return (short)ci.XFIndex;
            }
            return 0xF;
        }

        /**
         * Set the width for a given column in 1/256th of a Char width Units
         * @param column - the column number
         * @param width (in Units of 1/256th of a Char width)
         */
        public void SetColumnWidth(int column, int width)
        {
            if (width > 255 * 256) 
                   throw new ArgumentException("The maximum column width for an individual cell is 255 characters.");
            SetColumn(column, null, width, null, null, null);
        }

        /**
         * Get the hidden property for a given column.
         * @param column index
         * @see org.apache.poi.hssf.record.DefaultColWidthRecord
         * @see org.apache.poi.hssf.record.ColumnInfoRecord
         * @see #SetColumnHidden(short,bool)
         * @return whether the column is hidden or not.
         */

        public bool IsColumnHidden(int columnIndex)
        {
            ColumnInfoRecord cir = _columnInfos.FindColumnInfo(columnIndex);
            if (cir == null)
            {
                return false;
            }
            return cir.IsHidden;
        }

        /**
         * Get the hidden property for a given column.
         * @param column - the column number
         * @param hidden - whether the column is hidden or not
         */
        public void SetColumnHidden(int column, bool hidden)
        {
            SetColumn(column, null, null, null, hidden, null);
        }
        public void SetDefaultColumnStyle(int column, int styleIndex)
        {
            SetColumn(column, (short)styleIndex, null, null, null, null);
        }

        public void SetColumn(int column, int width, int level, bool hidden, bool collapsed)
        {
            _columnInfos.SetColumn(column, 0, width, level, hidden, collapsed);
        }

        public void SetColumn(int column, short? xfStyle, int? width, int? level, bool? hidden, bool? collapsed)
        {
            _columnInfos.SetColumn(column, xfStyle, width, level, hidden, collapsed);
        }

        private GutsRecord GetGutsRecord()
        {
            if (_gutsRecord == null)
            {
                GutsRecord result = CreateGuts();
                RecordOrderer.AddNewSheetRecord(records, result);
                _gutsRecord = result;
            }

            return _gutsRecord;
        }

        /**
         * Creates an outline Group for the specified columns.
         * @param fromColumn    Group from this column (inclusive)
         * @param toColumn      Group to this column (inclusive)
         * @param indent        if true the Group will be indented by one level,
         *                      if false indenting will be Removed by one level.
         */
        public void GroupColumnRange(int fromColumn, int toColumn, bool indent)
        {

            // Set the level for each column
            _columnInfos.GroupColumnRange(fromColumn, toColumn, indent);

            // Determine the maximum overall level
            int maxLevel = _columnInfos.MaxOutlineLevel;

            GutsRecord guts = GetGutsRecord();
            guts.ColLevelMax = (short)(maxLevel + 1);
            if (maxLevel == 0)
                guts.TopColGutter = ((short)0);
            else
                guts.TopColGutter = ((short)(29 + (12 * (maxLevel - 1))));
        }

        /**
         * Creates the Dimensions Record and Sets it to bogus values (you should Set this yourself
         * or let the high level API do it for you)
         * @see org.apache.poi.hssf.record.DimensionsRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a DimensionsRecord
         */

        private static DimensionsRecord CreateDimensions()
        {
            DimensionsRecord retval = new DimensionsRecord();

            retval.FirstCol = (short)0;
            retval.LastRow = 1;             // one more than it Is
            retval.FirstRow = 0;
            retval.LastCol = (short)1;   // one more than it Is
            return retval;
        }

        /**
         * Creates the WindowTwo Record and Sets it to:  
         * options        = 0x6b6 
         * toprow         = 0 
         * leftcol        = 0 
         * headercolor    = 0x40 
         * pagebreakzoom  = 0x0 
         * normalzoom     = 0x0 
         * @see org.apache.poi.hssf.record.WindowTwoRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a WindowTwoRecord
         */

        private static WindowTwoRecord CreateWindowTwo()
        {
            WindowTwoRecord retval = new WindowTwoRecord();

            retval.Options = ((short)0x6b6);
            retval.TopRow = ((short)0);
            retval.LeftCol = ((short)0);
            retval.HeaderColor = (0x40);
            retval.PageBreakZoom = ((short)0);
            retval.NormalZoom = ((short)0);
            return retval;
        }

        /// <summary>
        /// Creates the Selection record and Sets it to nothing selected
        /// </summary>
        /// <returns>record containing a SelectionRecord</returns>
        private static SelectionRecord CreateSelection()
        {
            return new SelectionRecord(0, 0);
        }

        /// <summary>
        /// Gets or sets the top row.
        /// </summary>
        /// <value>The top row.</value>
        public short TopRow
        {
            get
            {
                return (windowTwo == null) ? (short)0 : windowTwo.TopRow;
            }
            set
            {
                if (windowTwo != null)
                {
                    windowTwo.TopRow = (value);
                }
            }
        }

        /// <summary>
        /// Gets or sets the left col.
        /// </summary>
        /// <value>The left col.</value>
        public short LeftCol
        {
            get
            {
                return (windowTwo == null) ? (short)0 : windowTwo.LeftCol;
            }
            set
            {
                if (windowTwo != null)
                {
                    windowTwo.LeftCol = (value);
                }
            }
        }

        /// <summary>
        /// Sets the active cell.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        public void SetActiveCell(int row, int column)
        {
            this.SetActiveCellRange(row, row, column, column);
        }
        /// <summary>
        /// Sets the active cell range.
        /// </summary>
        /// <param name="firstRow">The firstrow.</param>
        /// <param name="lastRow">The lastrow.</param>
        /// <param name="firstColumn">The firstcolumn.</param>
        /// <param name="lastColumn">The lastcolumn.</param>
        public void SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            List<CellRangeAddress8Bit> cellranges = new List<CellRangeAddress8Bit>();
            cellranges.Add(new CellRangeAddress8Bit(firstRow, lastRow, firstColumn, lastColumn));
            this.SetActiveCellRange(cellranges, 0, firstRow, firstColumn);
        }
        /// <summary>
        /// Sets the active cell range.
        /// </summary>
        /// <param name="cellranges">The cellranges.</param>
        /// <param name="activeRange">The index of the active range.</param>
        /// <param name="activeRow">The active row in the active range</param>
        /// <param name="activeColumn">The active column in the active range</param>
        public void SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn)
        {
            this.selection.ActiveCellCol = activeColumn;
            this.selection.ActiveCellRow = activeRow;
            this.selection.ActiveCellRef = activeRange;
            this.selection.CellReferences = cellranges.ToArray();

        }

        /// <summary>
        /// Returns the active row
        /// </summary>
        /// <value>the active row index</value>
        /// @see org.apache.poi.hssf.record.SelectionRecord
        public int ActiveCellRow
        {
            get
            {
                if (selection == null)
                {
                    return 0;
                }
                return selection.ActiveCellRow;
            }
        }


        /// <summary>
        /// Gets the active cell col.
        /// </summary>
        /// <value>the active column index</value>
        /// @see org.apache.poi.hssf.record.SelectionRecord
        public int ActiveCellCol
        {
            get
            {
                if (selection == null)
                {
                    return 0;
                }
                return selection.ActiveCellCol;
            }
        }

        /// <summary>
        /// Creates the EOF record
        /// </summary>
        /// <returns>record containing a EOFRecord</returns>
        protected Record CreateEOF()
        {
            return new EOFRecord();
        }

        public List<RecordBase> Records
        {
            get { return records; }
        }

        /// <summary>
        /// Gets the gridset record for this sheet.
        /// </summary>
        /// <value>The gridset record.</value>
        public GridsetRecord GridsetRecord
        {
            get
            {
                return gridset;
            }
        }
        private GutsRecord GutsRecord
        {
            get
            {
                if (_gutsRecord == null)
                {
                    GutsRecord result = CreateGuts();
                    RecordOrderer.AddNewSheetRecord(records, result);
                    _gutsRecord = result;
                }

                return _gutsRecord;
            }
        }


        /// <summary>
        /// Returns the first occurance of a record matching a particular sid.
        /// </summary>
        /// <param name="sid">The sid.</param>
        /// <returns></returns>
        public Record FindFirstRecordBySid(short sid)
        {
            int ix = FindFirstRecordLocBySid(sid);
            if (ix < 0)
            {
                return null;
            }
            return (Record)records[ix];
        }

        /// <summary>
        /// Sets the SCL record or Creates it in the correct place if it does not
        /// already exist.
        /// </summary>
        /// <param name="sclRecord">The record to set.</param>
        public void SetSCLRecord(SCLRecord sclRecord)
        {
            int oldRecordLoc = FindFirstRecordLocBySid(SCLRecord.sid);
            if (oldRecordLoc == -1)
            {
                // Insert it after the window record
                int windowRecordLoc = FindFirstRecordLocBySid(WindowTwoRecord.sid);
                records.Insert(windowRecordLoc + 1, sclRecord);
            }
            else
            {
                records[oldRecordLoc] = sclRecord;
            }

        }
        /**
         * Finds the first occurance of a record matching a particular sid and
         * returns it's position.
         * @param sid   the sid to search for
         * @return  the record position of the matching record or -1 if no match
         *          is made.
         */
        public int FindFirstRecordLocBySid(short sid)
        {
            int max = records.Count;
            for (int i = 0; i < max; i++)
            {
                Object rb = records[i];
                if (!(rb is Record))
                {
                    continue;
                }
                Record record = (Record)rb;
                if (record.Sid == sid)
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Gets or sets the header.
        /// </summary>
        /// <value>the HeaderRecord.</value>
        public HeaderRecord Header
        {
            get { return header; }
            set { header = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is auto tab color.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is auto tab color; otherwise, <c>false</c>.
        /// </value>
        public bool IsAutoTabColor
        {
            get { return sheetext.IsAutoColor; }
            set
            {
                sheetext.IsAutoColor = value;
            }
        }

        public short TabColorIndex
        {
            get { return sheetext.TabColorIndex; }
            set
            {
                if ((value <= 0x08 || value >= 0x3F) && value != 0x7F)
                {
                    throw new ArgumentException("invalid color index");
                }
                sheetext.TabColorIndex = value;
            }
        }

        public WindowTwoRecord GetWindowTwo()
        {
            return windowTwo;
        }
        /// <summary>
        /// Gets or sets the footer.
        /// </summary>
        /// <value>FooterRecord for the sheet.</value>
        public FooterRecord Footer
        {
            get { return footer; }
            set { footer = value; }
        }


        /**
         * Returns the PrintSetupRecord.
         * @return PrintSetupRecord for the sheet.
         */
        public PrintSetupRecord PrintSetup
        {
            get { return printSetup; }
            set { printSetup = value; }
        }

        /**
 * @return <c>true</c> if gridlines are printed
 */
        public bool IsGridsPrinted
        {
            get
            {
                if (gridset == null)
                {
                    gridset = CreateGridset();
                    //Insert the newlycreated Gridset record at the end of the record (just before the EOF)
                    int loc = FindFirstRecordLocBySid(EOFRecord.sid);
                    records.Insert(loc, gridset);
                }
                return !gridset.Gridset;
            }
            set 
            {
                gridset.Gridset = !value;
            }
        }

        /**
         * Returns the PrintGridlinesRecord.
         * @return PrintGridlinesRecord for the sheet.
         */
        public PrintGridlinesRecord PrintGridlines
        {
            get { return printGridlines; }
            set { printGridlines = value; }
        }

        /**
         * Sets whether the sheet is selected
         * @param sel True to select the sheet, false otherwise.
         */
        public void SetSelected(bool sel)
        {
            windowTwo.IsSelected = (sel);
        }


        /**
         * Creates a split (freezepane). Any existing freezepane or split pane Is overwritten.
         * @param colSplit      Horizonatal position of split.
         * @param rowSplit      Vertical position of split.
         * @param topRow        Top row visible in bottom pane
         * @param leftmostColumn   Left column visible in right pane.
         */
        public void CreateFreezePane(int colSplit, int rowSplit, int topRow, int leftmostColumn)
        {
            int paneLoc = FindFirstRecordLocBySid(PaneRecord.sid);
            if (paneLoc != -1)
                records.RemoveAt(paneLoc);

            // If both colSplit and rowSplit are zero then the existing freeze pane is removed
            if (colSplit == 0 && rowSplit == 0)
            {
                windowTwo.FreezePanes = (false);
                windowTwo.FreezePanesNoSplit = (false);
                SelectionRecord sel = (SelectionRecord)FindFirstRecordBySid(SelectionRecord.sid);
                sel.Pane = (PaneInformation.PANE_UPPER_LEFT);
                return;
            }

            int loc = FindFirstRecordLocBySid(WindowTwoRecord.sid);
            PaneRecord pane = new PaneRecord();
            pane.X = ((short)colSplit);
            pane.Y = ((short)rowSplit);
            pane.TopRow = ((short)topRow);
            pane.LeftColumn = ((short)leftmostColumn);
            if (rowSplit == 0)
            {
                pane.TopRow = ((short)0);
                pane.ActivePane = ((short)1);
            }
            else if (colSplit == 0)
            {
                pane.LeftColumn = ((short)0);
                pane.ActivePane = ((short)2);
            }
            else
            {
                pane.ActivePane = ((short)0);
            }
            records.Insert(loc + 1, pane);

            windowTwo.FreezePanes = (true);
            windowTwo.FreezePanesNoSplit = (true);

            SelectionRecord sel1 = (SelectionRecord)FindFirstRecordBySid(SelectionRecord.sid);
            sel1.Pane = ((byte)pane.ActivePane);

        }

        /**
         * Creates a split pane. Any existing freezepane or split pane is overwritten.
         * @param xSplitPos      Horizonatal position of split (in 1/20th of a point).
         * @param ySplitPos      Vertical position of split (in 1/20th of a point).
         * @param topRow        Top row visible in bottom pane
         * @param leftmostColumn   Left column visible in right pane.
         * @param activePane    Active pane.  One of: PANE_LOWER_RIGHT,
         *                      PANE_UPPER_RIGHT, PANE_LOWER_LEFT, PANE_UPPER_LEFT
         * @see #PANE_LOWER_LEFT
         * @see #PANE_LOWER_RIGHT
         * @see #PANE_UPPER_LEFT
         * @see #PANE_UPPER_RIGHT
         */
        public void CreateSplitPane(int xSplitPos, int ySplitPos, int topRow, int leftmostColumn, NPOI.SS.UserModel.PanePosition activePane)
        {
            int paneLoc = FindFirstRecordLocBySid(PaneRecord.sid);
            if (paneLoc != -1)
                records.RemoveAt(paneLoc);

            int loc = FindFirstRecordLocBySid(WindowTwoRecord.sid);
            PaneRecord r = new PaneRecord();
            r.X = ((short)xSplitPos);
            r.Y = ((short)ySplitPos);
            r.TopRow = ((short)topRow);
            r.LeftColumn = ((short)leftmostColumn);
            r.ActivePane = ((short)activePane);
            records.Insert(loc + 1, r);

            windowTwo.FreezePanes = (false);
            windowTwo.FreezePanesNoSplit = (false);

            SelectionRecord sel = (SelectionRecord)FindFirstRecordBySid(SelectionRecord.sid);
            sel.Pane = (byte)NPOI.SS.UserModel.PanePosition.LowerRight;

        }

        /**
         * Returns the information regarding the currently configured pane (split or freeze).
         * @return null if no pane configured, or the pane information.
         */
        public PaneInformation PaneInformation
        {
            get
            {
                PaneRecord rec = (PaneRecord)FindFirstRecordBySid(PaneRecord.sid);
                if (rec == null)
                    return null;

                return new NPOI.SS.Util.PaneInformation(rec.X, rec.Y, rec.TopRow,
                                           rec.LeftColumn, (byte)rec.ActivePane, windowTwo.FreezePanes);
            }
        }

        public SelectionRecord Selection
        {
            get
            {
                return selection;
            }
            set { this.selection = value; }
        }
        /**
 * creates a Password record with password set to 00.
 */
        protected static PasswordRecord CreatePassword()
        {
            //if (log.check(POILogger.DEBUG))
            //    log.log(POILogger.DEBUG, "create password record with 00 password");
            PasswordRecord retval = new PasswordRecord((short)00);
            return retval;
        }
        /**
 * creates a Protect record with protect set to false.
 */
        protected ProtectRecord CreateProtect()
        {
            /*if (log.check(POILogger.DEBUG))
            {
                log.log(POILogger.DEBUG, "create protect record with protection disabled");
            }*/
            ProtectRecord retval = new ProtectRecord(false);
            return retval;
        }
        /**
         * Creates an ObjectProtect record with protect Set to false.
         * @see org.apache.poi.hssf.record.ObjectProtectRecord
         * @see org.apache.poi.hssf.record.Record
         * @return an ObjectProtectRecord
         */
        protected ObjectProtectRecord CreateObjectProtect()
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "Create protect record with protection disabled");
            ObjectProtectRecord retval = new ObjectProtectRecord();

            retval.Protect = (false);
            return retval;
        }

        /**
         * Creates a ScenarioProtect record with protect Set to false.
         * @see org.apache.poi.hssf.record.ScenarioProtectRecord
         * @see org.apache.poi.hssf.record.Record
         * @return a ScenarioProtectRecord
         */
        
        protected ScenarioProtectRecord CreateScenarioProtect()
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "Create protect record with protection disabled");
            ScenarioProtectRecord retval = new ScenarioProtectRecord();

            retval.Protect = (false);
            return retval;
        }
        
        //'Worksheet Protection Block'<br/>
        //Aggregate object is always present, but possibly empty.
        [NonSerialized]
        private WorksheetProtectionBlock _protectionBlock = new WorksheetProtectionBlock();
        
        public WorksheetProtectionBlock ProtectionBlock
        {
            get
            {
                return _protectionBlock;
            }
        }

        /**
         * Returns if gridlines are Displayed.
         * @return whether gridlines are Displayed
         */
        public bool DisplayGridlines
        {
            get
            {
                return windowTwo.DisplayGridlines;
            }
            set { windowTwo.DisplayGridlines = (value); }
        }

        /**
         * Returns if formulas are Displayed.
         * @return whether formulas are Displayed
         */
        public bool DisplayFormulas
        {
            get { return windowTwo.DisplayFormulas; }
            set { windowTwo.DisplayFormulas = (value); }
        }

        /**
         * Returns if RowColHeadings are Displayed.
         * @return whether RowColHeadings are Displayed
         */
        public bool DisplayRowColHeadings
        {
            get
            {
                return windowTwo.DisplayRowColHeadings;
            }
            set
            {
                windowTwo.DisplayRowColHeadings = (value);
            }
        }


        /**
         * @return whether an Uncalced record must be Inserted or not at generation
         */
        public bool IsUncalced
        {
            get { return _isUncalced; }
            set { this._isUncalced = value; }
        }

        /// <summary>
        /// Finds the DrawingRecord for our sheet, and  attaches it to the DrawingManager (which knows about
        ///  the overall DrawingGroup for our workbook).
        /// If requested, will Create a new DrawRecord if none currently exist
        /// </summary>
        /// <param name="drawingManager">The DrawingManager2 for our workbook</param>
        /// <param name="CreateIfMissing">Should one be Created if missing?</param>
        /// <returns>location of EscherAggregate record. if no EscherAggregate record is found return -1</returns>
        public int AggregateDrawingRecords(DrawingManager2 drawingManager, bool CreateIfMissing)
        {
            int loc = FindFirstRecordLocBySid(DrawingRecord.sid);
            bool noDrawingRecordsFound = (loc == -1);
            if (noDrawingRecordsFound)
            {
                if (!CreateIfMissing)
                {
                    // None found, and not allowed to Add in
                    return -1;
                }

                EscherAggregate aggregate = new EscherAggregate(true);
                loc = FindFirstRecordLocBySid(EscherAggregate.sid);
                if (loc == -1)
                {
                    loc = FindFirstRecordLocBySid(WindowTwoRecord.sid);
                }
                else
                {
                    Records.RemoveAt(loc);
                }
                Records.Insert(loc, aggregate);
                return loc;
            }
            EscherAggregate.CreateAggregate(records, loc);
            return loc;
        }

        /**
         * Perform any work necessary before the sheet is about to be Serialized.
         * For instance the escher aggregates size needs to be calculated before
         * serialization so that the dgg record (which occurs first) can be written.
         */
        public void Preserialize()
        {
            for (IEnumerator iterator = Records.GetEnumerator(); iterator.MoveNext(); )
            {
                RecordBase r = (RecordBase)iterator.Current;
                if (r is EscherAggregate)
                {
                    int size = r.RecordSize;   // Trigger flatterning of user model and corresponding update of dgg record.
                }
            }
        }

        /**
         * Shifts all the page breaks in the range "count" number of rows/columns
         * @param breaks The page record to be Shifted
         * @param start Starting "main" value to Shift breaks
         * @param stop Ending "main" value to Shift breaks
         * @param count number of Units (rows/columns) to Shift by
         */
        public void ShiftBreaks(PageBreakRecord breaks, short start, short stop, int count)
        {

            if (rowBreaks == null)
                return;
            IEnumerator iterator = breaks.GetBreaksEnumerator();
            IList ShiftedBreak = new ArrayList();
            while (iterator.MoveNext())
            {
                PageBreakRecord.Break breakItem = (PageBreakRecord.Break)iterator.Current;
                int breakLocation = breakItem.main;
                bool inStart = (breakLocation >= start);
                bool inEnd = (breakLocation <= stop);
                if (inStart && inEnd)
                    ShiftedBreak.Add(breakItem);
            }

            iterator = ShiftedBreak.GetEnumerator();
            while (iterator.MoveNext())
            {
                PageBreakRecord.Break breakItem = (PageBreakRecord.Break)iterator.Current;
                breaks.RemoveBreak(breakItem.main);
                breaks.AddBreak(breakItem.main + count, breakItem.subFrom, breakItem.subTo);
            }
        }


        /**
         * Shifts the horizontal page breaks for the indicated count
         * @param startingRow
         * @param endingRow
         * @param count
         */
        public void ShiftRowBreaks(int startingRow, int endingRow, int count)
        {
            ShiftBreaks(rowBreaks, (short)startingRow, (short)endingRow, (short)count);
        }

        /**
         * Shifts the vertical page breaks for the indicated count
         * @param startingCol
         * @param endingCol
         * @param count
         */
        public void ShiftColumnBreaks(short startingCol, short endingCol, short count)
        {
            ShiftBreaks(colBreaks, startingCol, endingCol, count);
        }


        public void SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            if (collapsed)
            {
                _columnInfos.CollapseColumn(columnNumber);
            }
            else
            {
                _columnInfos.ExpandColumn(columnNumber);
            }
        }
        public RowRecordsAggregate RowsAggregate
        {
            get
            {
                return _rowsAggregate;
            }
        }
        /**
 * Updates formulas in cells and conditional formats due to moving of cells
 * @param externSheetIndex the externSheet index of this sheet
 */
        public void UpdateFormulasAfterCellShift(FormulaShifter shifter, int externSheetIndex)
        {
            RowsAggregate.UpdateFormulasAfterRowShift(shifter, externSheetIndex);
            if (condFormatting != null)
            {
                ConditionalFormattingTable.UpdateFormulasAfterCellShift(shifter, externSheetIndex);
            }
            // TODO - adjust data validations
        }

        public ConditionalFormattingTable ConditionalFormattingTable
        {
            get
            {
                if (condFormatting == null)
                {
                    condFormatting = new ConditionalFormattingTable();
                    RecordOrderer.AddNewSheetRecord(records, condFormatting);
                }
                return condFormatting;
            }
        }

        public void VisitContainedRecords(RecordVisitor rv, int offset)
        {

            PositionTrackingVisitor ptv = new PositionTrackingVisitor(rv, offset);

            bool haveSerializedIndex = false;

            int sheetOffset = offset;
            for (int k = 0; k < records.Count; k++)
            {
                RecordBase record = records[k];

                if (record is RecordAggregate)
                {
                    RecordAggregate agg = (RecordAggregate)record;
                    agg.VisitContainedRecords(ptv);
                    sheetOffset += agg.RecordSize;
                }
                else
                {
                    if (record is DefaultColWidthRecord)
                    {
                        ((DefaultColWidthRecord)record).offsetForFilePointer = sheetOffset;
                    }
                    ptv.VisitRecord((Record)record);
                    sheetOffset += record.RecordSize;
                }

                // If the BOF record was just serialized then add the IndexRecord
                if (record is BOFRecord)
                {
                    if (!haveSerializedIndex)
                    {
                        haveSerializedIndex = true;
                        // Add an optional UncalcedRecord. However, we should add
                        //  it in only the once, after the sheet's own BOFRecord.
                        // If there are diagrams, they have their own BOFRecords,
                        //  and one shouldn't go in after that!
                        if (_isUncalced)
                        {
                            UncalcedRecord rec = new UncalcedRecord();
                            ptv.VisitRecord(rec);
                            sheetOffset += rec.RecordSize;
                        }
                        //Can there be more than one BOF for a sheet? If not then we can
                        //remove this guard. So be safe it is left here.
                        if (_rowsAggregate != null)
                        {
                            // find forward distance to first RowRecord
                            int initRecsSize = GetSizeOfInitialSheetRecords(k);
                            int currentPos = ptv.Position;
                            IndexRecord indexRecord = _rowsAggregate.CreateIndexRecord(currentPos, initRecsSize, 0);
                            ptv.VisitRecord(indexRecord);
                            sheetOffset += indexRecord.RecordSize;
                        }
                    }
                }
            }
        }

        /**
 * 'initial sheet records' are between INDEX and the 'Row Blocks'
 * @param bofRecordIndex index of record after which INDEX record is to be placed
 * @return count of bytes from end of INDEX record to first ROW record.
 */
        private int GetSizeOfInitialSheetRecords(int bofRecordIndex)
        {

            int result = 0;
            // start just after BOF record (INDEX is not present in this list)
            for (int j = bofRecordIndex + 1; j < records.Count; j++)
            {
                RecordBase tmpRec = records[j];
                if (tmpRec is RowRecordsAggregate)
                {
                    break;
                }
                result += tmpRec.RecordSize;
            }
            if (_isUncalced)
            {
                result += UncalcedRecord.StaticRecordSize;
            }
            return result;
        }

        public void GroupRowRange(int fromRow, int toRow, bool indent)
        {
            for (int rowNum = fromRow; rowNum <= toRow; rowNum++)
            {
                RowRecord row = GetRow(rowNum);
                if (row == null)
                {
                    row = CreateRow(rowNum);
                    AddRow(row);
                }
                int level = row.OutlineLevel;
                if (indent) level++; else level--;
                level = Math.Max(0, level);
                level = Math.Min(7, level);
                row.OutlineLevel = ((short)(level));
            }

            RecalcRowGutter();
        }

        private void RecalcRowGutter()
        {
            int maxLevel = 0;
            IEnumerator iterator = _rowsAggregate.GetEnumerator();
            while (iterator.MoveNext())
            {
                RowRecord rowRecord = (RowRecord)iterator.Current;
                maxLevel = Math.Max(rowRecord.OutlineLevel, maxLevel);
            }

            // Grab the guts record, Adding if needed
            GutsRecord guts = GetGutsRecord();
            if (guts == null)
            {
                guts = new GutsRecord();
                records.Add(guts);
            }
            // Set the levels onto it
            guts.RowLevelMax = ((short)(maxLevel + 1));
            guts.LeftRowGutter = (short)(29 + 12 * maxLevel);
        }
        public DataValidityTable GetOrCreateDataValidityTable()
        {
            if (_dataValidityTable == null)
            {
                _dataValidityTable = new DataValidityTable();
                RecordOrderer.AddNewSheetRecord(records, _dataValidityTable);
            }
            return _dataValidityTable;
        }
        /**
         * Get the {@link NoteRecord}s (related to cell comments) for this sheet
         * @return never <code>null</code>, typically empty array
         */
        public NoteRecord[] GetNoteRecords()
        {
            List<NoteRecord> temp = new List<NoteRecord>();
            for (int i = records.Count - 1; i >= 0; i--)
            {
                RecordBase rec = records[i];
                if (rec is NoteRecord)
                {
                    temp.Add((NoteRecord)rec);
                }
            }
            if (temp.Count < 1)
            {
                return NoteRecord.EMPTY_ARRAY;
            }
            NoteRecord[] result = new NoteRecord[temp.Count];
            result = temp.ToArray();
            return result;
        }

        public int GetColumnOutlineLevel(int columnIndex)
        {
            return _columnInfos.GetOutlineLevel(columnIndex);
        }

    }

    public class UnsupportedBOFType : RecordFormatException
    {
        private BOFRecordType type;
        public UnsupportedBOFType(BOFRecordType type)
            : base("BOF not of a supported type, found " + type)
        {
            ;
            this.type = type;
        }

        public BOFRecordType Type
        {
            get
            {
                return type;
            }
        }
    }
}
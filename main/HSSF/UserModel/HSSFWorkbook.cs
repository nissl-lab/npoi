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

namespace NPOI.HSSF.UserModel
{

    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Reflection;
    using System.Security.Cryptography;
    using System.Text;
    using NPOI.DDF;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Udf;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;


    /// <summary>
    /// High level representation of a workbook.  This is the first object most users
    /// will construct whether they are reading or writing a workbook.  It is also the
    /// top level object for creating new sheets/etc.
    /// </summary>
    /// @author  Andrew C. Oliver (acoliver at apache dot org)
    /// @author  Glen Stampoultzis (glens at apache.org)
    /// @author  Shawn Laubach (slaubach at apache dot org)
    [Serializable]
    public class HSSFWorkbook : POIDocument, IWorkbook
    {
        //private static int DEBUG = POILogger.DEBUG;

        /**
         * The maximum number of cell styles in a .xls workbook.
         * The 'official' limit is 4,000, but POI allows a slightly larger number.
         * This extra delta takes into account built-in styles that are automatically
         * created for new workbooks
         *
         * See http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP005199291.aspx
         */
        private const int MAX_STYLES = 4030;

        /**
         * used for compile-time performance/memory optimization.  This determines the
         * initial capacity for the sheet collection.  Its currently Set to 3.
         * Changing it in this release will decrease performance
         * since you're never allowed to have more or less than three sheets!
         */

        public const int INITIAL_CAPACITY = 3;

        /**
         * this Is the reference to the low level Workbook object
         */

        private InternalWorkbook workbook;

        /**
         * this holds the HSSFSheet objects attached to this workbook
         */

        protected List<HSSFSheet> _sheets;

        /**
         * this holds the HSSFName objects attached to this workbook
         */

        private List<HSSFName> names;

        /**
         * holds whether or not to preserve other nodes in the POIFS.  Used
         * for macros and embedded objects.
         */
        private bool preserveNodes;

        /**
         * Used to keep track of the data formatter so that all
         * CreateDataFormatter calls return the same one for a given
         * book.  This Ensures that updates from one places Is visible
         * someplace else.
         */
        private HSSFDataFormat formatter;

        /**
         * this holds the HSSFFont objects attached to this workbook.
         * We only create these from the low level records as required.
         */
        private Hashtable fonts;



        //private static POILogger log = POILogFactory.GetLogger(typeof(HSSFWorkbook));

        public NPOI.SS.UserModel.ICreationHelper GetCreationHelper()
        {
            return new HSSFCreationHelper(this);
        }

        /// <summary>
        /// Totals the sizes of all sheet records and eventually serializes them
        /// </summary>
        private class SheetRecordCollector : NPOI.HSSF.Record.Aggregates.RecordVisitor,IDisposable
        {

            private ArrayList _list;
            private int _totalSize;

            public SheetRecordCollector()
            {
                _totalSize = 0;
                _list = new ArrayList(128);
            }
            public int TotalSize
            {
                get
                {
                    return _totalSize;
                }
            }
            public void VisitRecord(Record r)
            {
                _list.Add(r);
                _totalSize += r.RecordSize;
            }
            public int Serialize(int offset, byte[] data)
            {
                int result = 0;
                int nRecs = _list.Count;
                for (int i = 0; i < nRecs; i++)
                {
                    Record rec = (Record)_list[i];
                    result += rec.Serialize(offset + result, data);
                }
                return result;
            }
            public void Dispose()
            {
                //this._list = null;
            }

        }
        public static HSSFWorkbook Create(InternalWorkbook book)
        {
            return new HSSFWorkbook(book);
        }

        /// <summary>
        /// Creates new HSSFWorkbook from scratch (start here!)
        /// </summary>
        public HSSFWorkbook()
            : this(InternalWorkbook.CreateWorkbook())
        {

        }

        public HSSFWorkbook(InternalWorkbook book)
            : base((DirectoryNode)null)
        {

            workbook = book;
            _sheets = new List<HSSFSheet>(INITIAL_CAPACITY);
            names = new List<HSSFName>(INITIAL_CAPACITY);
        }
        /**
         * Companion to HSSFWorkbook(POIFSFileSystem), this constructs the 
         *  POI filesystem around your inputstream, including all nodes.
         * This calls {@link #HSSFWorkbook(InputStream, boolean)} with
         *  preserve nodes set to true. 
         *
         * @see #HSSFWorkbook(InputStream, boolean)
         * @see #HSSFWorkbook(POIFSFileSystem)
         * @see org.apache.poi.poifs.filesystem.POIFSFileSystem
         * @exception IOException if the stream cannot be read
         */
        public HSSFWorkbook(POIFSFileSystem fs)
            : this(fs, true)
        {

        }

        /// <summary>
        /// given a POI POIFSFileSystem object, Read in its Workbook and populate the high and
        /// low level models.  If you're Reading in a workbook...start here.
        /// </summary>
        /// <param name="fs">the POI filesystem that Contains the Workbook stream.</param>
        /// <param name="preserveNodes">whether to preseve other nodes, such as
        /// macros.  This takes more memory, so only say yes if you
        /// need to. If Set, will store all of the POIFSFileSystem
        /// in memory</param>
        public HSSFWorkbook(POIFSFileSystem fs, bool preserveNodes)
            : this(fs.Root, fs, preserveNodes)
        {

        }

        /**
         * Normally, the Workbook will be in a POIFS Stream
         * called "Workbook". However, some weird XLS generators use "WORKBOOK"
         */
        private static readonly string[] WORKBOOK_DIR_ENTRY_NAMES = {
            "Workbook", // as per BIFF8 spec
            "WORKBOOK", // Typically from third party programs
            "BOOK",     // Typically odd Crystal Reports exports
        };


        private static String GetWorkbookDirEntryName(DirectoryNode directory)
        {

            String[] potentialNames = WORKBOOK_DIR_ENTRY_NAMES;
            for (int i = 0; i < potentialNames.Length; i++)
            {
                String wbName = potentialNames[i];
                try
                {
                    directory.GetEntry(wbName);
                    return wbName;
                }
                catch (FileNotFoundException)
                {
                    // continue - to try other options
                }
            }
            // check for an encrypted .xlsx file - they get OLE2 wrapped
            try
            {
                directory.GetEntry("EncryptedPackage");
                throw new EncryptedDocumentException("The supplied spreadsheet seems to be an Encrypted .xlsx file. " +
                        "It must be decrypted before use by XSSF, it cannot be used by HSSF");
            }
            catch (FileNotFoundException)
            {
                // fall through
            }
            // Check for previous version of file format
            try
            {
                directory.GetEntry("Book");
                throw new OldExcelFormatException("The supplied spreadsheet seems to be Excel 5.0/7.0 (BIFF5) format. "
                        + "POI only supports BIFF8 format (from Excel versions 97/2000/XP/2003)");
            }
            catch (FileNotFoundException)
            {
                // fall through
            }

            throw new ArgumentException("The supplied POIFSFileSystem does not contain a BIFF8 'Workbook' entry. "
                + "Is it really an excel file?");
        }

        /// <summary>
        /// given a POI POIFSFileSystem object, and a specific directory
        /// within it, Read in its Workbook and populate the high and
        /// low level models.  If you're Reading in a workbook...start here.
        /// </summary>
        /// <param name="directory">the POI filesystem directory to Process from</param>
        /// <param name="fs">the POI filesystem that Contains the Workbook stream.</param>
        /// <param name="preserveNodes">whether to preseve other nodes, such as
        /// macros.  This takes more memory, so only say yes if you
        /// need to. If Set, will store all of the POIFSFileSystem
        /// in memory</param>
        public HSSFWorkbook(DirectoryNode directory, POIFSFileSystem fs, bool preserveNodes)
            : this(directory, preserveNodes)
        {
        }
            /**
     * given a POI POIFSFileSystem object, and a specific directory
     *  within it, read in its Workbook and populate the high and
     *  low level models.  If you're reading in a workbook...start here.
     *
     * @param directory the POI filesystem directory to process from
     * @param preserveNodes whether to preseve other nodes, such as
     *        macros.  This takes more memory, so only say yes if you
     *        need to. If set, will store all of the POIFSFileSystem
     *        in memory
     * @see org.apache.poi.poifs.filesystem.POIFSFileSystem
     * @exception IOException if the stream cannot be read
     */
        public HSSFWorkbook(DirectoryNode directory, bool preserveNodes):base(directory)
        {

            String workbookName = GetWorkbookDirEntryName(directory);

            this.preserveNodes = preserveNodes;

            // If we're not preserving nodes, don't track the
            //  POIFS any more
            if (!preserveNodes)
            {
                this.directory = null;
            }

            _sheets = new List<HSSFSheet>(INITIAL_CAPACITY);
            names = new List<HSSFName>(INITIAL_CAPACITY);

            // Grab the data from the workbook stream, however
            //  it happens to be spelled.
            Stream stream = directory.CreatePOIFSDocumentReader(workbookName);


            List<Record> records = RecordFactory.CreateRecords(stream);

            workbook = InternalWorkbook.CreateWorkbook(records);
            SetPropertiesFromWorkbook(workbook);
            int recOffset = workbook.NumRecords;

            // Convert all LabelRecord records to LabelSSTRecord
            ConvertLabelRecords(records, recOffset);
            RecordStream rs = new RecordStream(records, recOffset);
            while (rs.HasNext())
            {
                try
                {
                    InternalSheet sheet = InternalSheet.CreateSheet(rs);
                    _sheets.Add(new HSSFSheet(this, sheet));
                }
                catch (UnsupportedBOFType eb)
                {
                    // Hopefully there's a supported one after this!
                    Console.WriteLine("Unsupported BOF found of type " + eb.Type);
                }
            }

            for (int i = 0; i < workbook.NumNames; ++i)
            {
                NameRecord nameRecord = workbook.GetNameRecord(i);
                HSSFName name = new HSSFName(this, workbook.GetNameRecord(i), workbook.GetNameCommentRecord(nameRecord));
                names.Add(name);
            }
        }

        public HSSFWorkbook(Stream s)
            : this(s, true)
        {

        }

        /**
         * Companion to HSSFWorkbook(POIFSFileSystem), this constructs the POI filesystem around your
         * inputstream.
         *
         * @param s  the POI filesystem that Contains the Workbook stream.
         * @param preserveNodes whether to preseve other nodes, such as
         *        macros.  This takes more memory, so only say yes if you
         *        need to.
         * @see org.apache.poi.poifs.filesystem.POIFSFileSystem
         * @see #HSSFWorkbook(POIFSFileSystem)
         * @exception IOException if the stream cannot be Read
         */

        public HSSFWorkbook(Stream s, bool preserveNodes)
            : this(new POIFSFileSystem(s), preserveNodes)
        {

        }

        /**
         * used internally to Set the workbook properties.
         */

        private void SetPropertiesFromWorkbook(InternalWorkbook book)
        {
            this.workbook = book;

            // none currently
        }

        /// <summary>
        /// This is basically a kludge to deal with the now obsolete Label records.  If
        /// you have to read in a sheet that contains Label records, be aware that the rest
        /// of the API doesn't deal with them, the low level structure only provides Read-only
        /// semi-immutable structures (the Sets are there for interface conformance with NO
        /// impelmentation).  In short, you need to call this function passing it a reference
        /// to the Workbook object.  All labels will be converted to LabelSST records and their
        /// contained strings will be written to the Shared String tabel (SSTRecord) within
        /// the Workbook.
        /// </summary>
        /// <param name="records">The records.</param>
        /// <param name="offset">The offset.</param>
        private void ConvertLabelRecords(IList records, int offset)
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "ConvertLabelRecords called");
            for (int k = offset; k < records.Count; k++)
            {
                Record rec = (Record)records[k];

                if (rec.Sid == LabelRecord.sid)
                {
                    LabelRecord oldrec = (LabelRecord)rec;

                    records.RemoveAt(k);
                    LabelSSTRecord newrec = new LabelSSTRecord();
                    int stringid =
                        workbook.AddSSTString(new UnicodeString(oldrec.Value));

                    newrec.Row = (oldrec.Row);
                    newrec.Column = (oldrec.Column);
                    newrec.XFIndex = (oldrec.XFIndex);
                    newrec.SSTIndex = (stringid);
                    records.Insert(k, newrec);
                }
            }
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(POILogger.DEBUG, "ConvertLabelRecords exit");
        }
        [NonSerialized]
        private NPOI.SS.UserModel.MissingCellPolicy missingCellPolicy = NPOI.SS.UserModel.MissingCellPolicy.RETURN_NULL_AND_BLANK;
        /// <summary>
        /// Retrieves the current policy on what to do when
        /// getting missing or blank cells from a row.
        /// The default is to return blank and null cells.
        /// </summary>
        /// <value>The missing cell policy.</value>
        public NPOI.SS.UserModel.MissingCellPolicy MissingCellPolicy
        {
            get { return missingCellPolicy; }
            set { this.missingCellPolicy = value; }
        }

        /// <summary>
        /// Sets the order of appearance for a given sheet.
        /// </summary>
        /// <param name="sheetname">the name of the sheet to reorder</param>
        /// <param name="pos">the position that we want to Insert the sheet into (0 based)</param>
        public void SetSheetOrder(String sheetname, int pos)
        {
            int oldSheetIndex = GetSheetIndex(sheetname);
            HSSFSheet sheet = (HSSFSheet)this.GetSheet(sheetname);
            _sheets.RemoveAt(oldSheetIndex);
            _sheets.Insert(pos, sheet);
            workbook.SetSheetOrder(sheetname, pos);

            FormulaShifter shifter = FormulaShifter.CreateForSheetShift(oldSheetIndex, pos);
            foreach (HSSFSheet sheet1 in _sheets)
            {
                sheet1.Sheet.UpdateFormulasAfterCellShift(shifter, /* not used */ -1);
            }

            workbook.UpdateNamesAfterCellShift(shifter);

            // adjust active sheet if necessary
            int active = ActiveSheetIndex;
            if (active == oldSheetIndex)
            {
                // moved sheet was the active one
                SetActiveSheet(pos);
            }
            else if ((active < oldSheetIndex && active < pos) ||
                  (active > oldSheetIndex && active > pos))
            {
                // not affected
            }
            else if (pos > oldSheetIndex)
            {
                // moved sheet was below before and is above now => active is one less
                SetActiveSheet(active - 1);
            }
            else
            {
                // remaining case: moved sheet was higher than active before and is lower now => active is one more
                SetActiveSheet(active + 1);
            }
        }

        /// <summary>
        /// Validates the index of the sheet.
        /// </summary>
        /// <param name="index">The index.</param>
        private void ValidateSheetIndex(int index)
        {
            int lastSheetIx = _sheets.Count - 1;
            if (index < 0 || index > lastSheetIx)
            {
                String range = "(0.." + lastSheetIx + ")";
                if (lastSheetIx == -1)
                {
                    range = "(no sheets)";
                }
                throw new ArgumentException("Sheet index ("
                        + index + ") is out of range " + range);
            }
        }
        /** Test only. Do not use */
        public void InsertChartRecord()
        {
            int loc = workbook.FindFirstRecordLocBySid(SSTRecord.sid);
            byte[] data = {
               (byte)0x0F, (byte)0x00, (byte)0x00, (byte)0xF0, (byte)0x52,
               (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
               (byte)0x06, (byte)0xF0, (byte)0x18, (byte)0x00, (byte)0x00,
               (byte)0x00, (byte)0x01, (byte)0x08, (byte)0x00, (byte)0x00,
               (byte)0x02, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x02,
               (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x01, (byte)0x00,
               (byte)0x00, (byte)0x00, (byte)0x01, (byte)0x00, (byte)0x00,
               (byte)0x00, (byte)0x03, (byte)0x00, (byte)0x00, (byte)0x00,
               (byte)0x33, (byte)0x00, (byte)0x0B, (byte)0xF0, (byte)0x12,
               (byte)0x00, (byte)0x00, (byte)0x00, (byte)0xBF, (byte)0x00,
               (byte)0x08, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x81,
               (byte)0x01, (byte)0x09, (byte)0x00, (byte)0x00, (byte)0x08,
               (byte)0xC0, (byte)0x01, (byte)0x40, (byte)0x00, (byte)0x00,
               (byte)0x08, (byte)0x40, (byte)0x00, (byte)0x1E, (byte)0xF1,
               (byte)0x10, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x0D,
               (byte)0x00, (byte)0x00, (byte)0x08, (byte)0x0C, (byte)0x00,
               (byte)0x00, (byte)0x08, (byte)0x17, (byte)0x00, (byte)0x00,
               (byte)0x08, (byte)0xF7, (byte)0x00, (byte)0x00, (byte)0x10,
            };
            UnknownRecord r = new UnknownRecord((short)0x00EB, data);
            workbook.Records.Insert(loc, r);
        }
        /// <summary>
        /// Selects a single sheet. This may be different to
        /// the 'active' sheet (which Is the sheet with focus).
        /// </summary>
        /// <param name="index">The index.</param>
        public void SetSelectedTab(int index)
        {

            ValidateSheetIndex(index);
            int nSheets = _sheets.Count;
            for (int i = 0; i < nSheets; i++)
            {
                GetSheetAt(i).IsSelected = (i == index);
            }
            workbook.WindowOne.NumSelectedTabs = ((short)1);
        }
        /// <summary>
        /// Sets the selected tabs.
        /// </summary>
        /// <param name="indexes">The indexes.</param>
        public void SetSelectedTabs(int[] indexes)
        {

            for (int i = 0; i < indexes.Length; i++)
            {
                ValidateSheetIndex(indexes[i]);
            }
            int nSheets = _sheets.Count;
            for (int i = 0; i < nSheets; i++)
            {
                bool bSelect = false;
                for (int j = 0; j < indexes.Length; j++)
                {
                    if (indexes[j] == i)
                    {
                        bSelect = true;
                        break;
                    }

                }
                GetSheetAt(i).IsSelected = (bSelect);
            }
            workbook.WindowOne.NumSelectedTabs = ((short)indexes.Length);
        }

        /// <summary>
        /// Gets the tab whose data is actually seen when the sheet is opened.
        /// This may be different from the "selected sheet" since excel seems to
        /// allow you to show the data of one sheet when another Is seen "selected"
        /// in the tabs (at the bottom).
        /// </summary>
        public int ActiveSheetIndex
        {
            get
            {
                return workbook.WindowOne.ActiveSheetIndex;
            }
        }

        /// <summary>
        /// Sets the tab whose data is actually seen when the sheet is opened.
        /// This may be different from the "selected sheet" since excel seems to
        /// allow you to show the data of one sheet when another Is seen "selected"
        /// in the tabs (at the bottom).
        /// <param name="index">The sheet number(0 based).</param>
        /// </summary>
        public void SetActiveSheet(int index)
        {

            ValidateSheetIndex(index);
            int nSheets = _sheets.Count;
            for (int i = 0; i < nSheets; i++)
            {
                GetSheetAt(i).SetActive(i == index);
            }
            workbook.WindowOne.ActiveSheetIndex = (index);
        }

        /// <summary>
        /// Gets or sets the first tab that is displayed in the list of tabs
        /// in excel.
        /// </summary>
        public int FirstVisibleTab
        {
            get { return workbook.WindowOne.FirstVisibleTab; }
            set { workbook.WindowOne.FirstVisibleTab = value;}
        }
        [Obsolete("Misleading name - use GetFirstVisibleTab() ")]
        public short DisplayedTab
        {
            get { return (short)FirstVisibleTab; }
        }

        /**
         * @deprecated POI will now properly handle Unicode strings without
         * forceing an encoding
         */
        public const byte ENCODING_COMPRESSED_UNICODE = 0;
        /**
         * @deprecated POI will now properly handle Unicode strings without
         * forceing an encoding
         */
        public const byte ENCODING_UTF_16 = 1;


        /// <summary>
        /// Set the sheet name.
        /// </summary>
        /// <param name="sheetIx">The sheet number(0 based).</param>
        /// <param name="name">The name.</param>
        public void SetSheetName(int sheetIx, String name)
        {
            if (name == null)
            {
                throw new ArgumentException("sheetName must not be null");
            }

            if (workbook.ContainsSheetName(name, sheetIx))
            {
                throw new ArgumentException("The workbook already contains a sheet with this name");
            }
            ValidateSheetIndex(sheetIx);
            workbook.SetSheetName(sheetIx, name);
        }

        /// <summary>
        /// Get the sheet name
        /// </summary>
        /// <param name="sheetIx">The sheet index.</param>
        /// <returns>Sheet name</returns>
        public String GetSheetName(int sheetIx)
        {
            ValidateSheetIndex(sheetIx);
            return workbook.GetSheetName(sheetIx);
        }

        /// <summary>
        /// Check whether a sheet is hidden
        /// </summary>
        /// <param name="sheetIx">The sheet index.</param>
        /// <returns>
        /// 	<c>true</c> if sheet is hidden; otherwise, <c>false</c>.
        /// </returns>
        public bool IsSheetHidden(int sheetIx)
        {
            ValidateSheetIndex(sheetIx);
            return workbook.IsSheetHidden(sheetIx);
        }
        /// <summary>
        /// Check whether a sheet is very hidden.
        /// This is different from the normal
        /// hidden status
        /// </summary>
        /// <param name="sheetIx">The sheet index.</param>
        /// <returns>
        /// 	<c>true</c> if sheet is very hidden; otherwise, <c>false</c>.
        /// </returns>
        public bool IsSheetVeryHidden(int sheetIx)
        {
            ValidateSheetIndex(sheetIx);
            return workbook.IsSheetVeryHidden(sheetIx);
        }
        /// <summary>
        /// Hide or Unhide a sheet
        /// </summary>
        /// <param name="sheetIx">The sheet index</param>
        /// <param name="hidden">True to mark the sheet as hidden, false otherwise</param>
        public void SetSheetHidden(int sheetIx, SheetState hidden)
        {
            ValidateSheetIndex(sheetIx);
            WorkbookUtil.ValidateSheetState(hidden);
            workbook.SetSheetHidden(sheetIx, (int)hidden);
        }
        /// <summary>
        /// Hide or unhide a sheet.
        /// </summary>
        /// <param name="sheetIx">The sheet number</param>
        /// <param name="hidden">0 for not hidden, 1 for hidden, 2 for very hidden</param>
        public void SetSheetHidden(int sheetIx, int hidden)
        {
            ValidateSheetIndex(sheetIx);
            workbook.SetSheetHidden(sheetIx, hidden);
        }
        public void SetSheetHidden(int sheetIx, bool hidden)
        {
            ValidateSheetIndex(sheetIx);
            workbook.SetSheetHidden(sheetIx, hidden);
        }
        /// <summary>
        /// Returns the index of the sheet by his name
        /// </summary>
        /// <param name="name">the sheet name</param>
        /// <returns>index of the sheet (0 based)</returns>
        public int GetSheetIndex(String name)
        {
            return workbook.GetSheetIndex(name);
        }

        /// <summary>
        /// Returns the index of the given sheet
        /// </summary>
        /// <param name="sheet">the sheet to look up</param>
        /// <returns>index of the sheet (0 based).-1
        ///  if not found </returns>
        public int GetSheetIndex(ISheet sheet)
        {
            for (int i = 0; i < _sheets.Count; i++)
            {
                if (_sheets[i] == sheet)
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns the external sheet index of the sheet
        /// with the given internal index, creating one
        /// if needed.
        /// Used by some of the more obscure formula and
        /// named range things.
        /// </summary>
        /// <param name="internalSheetIndex">Index of the internal sheet.</param>
        /// <returns></returns>
        public int GetExternalSheetIndex(int internalSheetIndex)
        {
            return workbook.CheckExternSheet(internalSheetIndex);
        }

        /// <summary>
        /// Create an HSSFSheet for this HSSFWorkbook, Adds it to the sheets and returns
        /// the high level representation.  Use this to Create new sheets.
        /// </summary>
        /// <returns>HSSFSheet representing the new sheet.</returns>
        public ISheet CreateSheet()
        {
            HSSFSheet sheet = new HSSFSheet(this);

            _sheets.Add(sheet);
            workbook.SetSheetName(_sheets.Count - 1, "Sheet" + (_sheets.Count - 1));
            bool IsOnlySheet = _sheets.Count == 1;
            sheet.IsSelected = (IsOnlySheet);
            sheet.IsActive = (IsOnlySheet);
            return sheet;
        }

        /// <summary>
        /// Create an HSSFSheet from an existing sheet in the HSSFWorkbook.
        /// </summary>
        /// <param name="sheetIndex">the sheet index</param>
        /// <returns>HSSFSheet representing the Cloned sheet.</returns>
        public ISheet CloneSheet(int sheetIndex)
        {
            ValidateSheetIndex(sheetIndex);
            HSSFSheet srcSheet = (HSSFSheet)_sheets[sheetIndex];
            String srcName = workbook.GetSheetName(sheetIndex);
            ISheet clonedSheet = srcSheet.CloneSheet(this);
            clonedSheet.IsSelected = (false);
            clonedSheet.IsActive = (false);

            String name = GetUniqueSheetName(srcName);
            int newSheetIndex = _sheets.Count;
            _sheets.Add((HSSFSheet)clonedSheet);
            workbook.SetSheetName(newSheetIndex, name);

            // Check this sheet has an autofilter, (which has a built-in NameRecord at workbook level)
            int filterDbNameIndex = FindExistingBuiltinNameRecordIdx(sheetIndex, NameRecord.BUILTIN_FILTER_DB);
            if (filterDbNameIndex != -1)
            {
                NameRecord newNameRecord = workbook.CloneFilter(filterDbNameIndex, newSheetIndex);
                HSSFName newName = new HSSFName(this, newNameRecord);
                names.Add(newName);
            }
            // TODO - maybe same logic required for other/all built-in name records
            //workbook.CloneDrawings(((HSSFSheet)clonedSheet).Sheet);
            return clonedSheet;
        }
        /// <summary>
        /// Gets the name of the unique sheet.
        /// </summary>
        /// <param name="srcName">Name of the SRC.</param>
        /// <returns></returns>
        private String GetUniqueSheetName(String srcName)
        {
            int uniqueIndex = 2;
            String baseName = srcName;
            int bracketPos = srcName.LastIndexOf('(');
            if (bracketPos > 0 && srcName.EndsWith(")", StringComparison.Ordinal))
            {
                String suffix = srcName.Substring(bracketPos + 1, srcName.Length - bracketPos - 2);
                try
                {
                    uniqueIndex = Int32.Parse(suffix.Trim(), CultureInfo.InvariantCulture);
                    uniqueIndex++;
                    baseName = srcName.Substring(0, bracketPos).Trim();
                }
                catch (FormatException)
                {
                    // contents of brackets not numeric
                }
            }
            while (true)
            {
                // Try and find the next sheet name that is unique
                String index = (uniqueIndex++).ToString(CultureInfo.CurrentCulture);
                String name;
                if (baseName.Length + index.Length + 2 < 31)
                {
                    name = baseName + " (" + index + ")";
                }
                else
                {
                    name = baseName.Substring(0, 31 - index.Length - 2) + "(" + index + ")";
                }

                //If the sheet name is unique, then set it otherwise move on to the next number.
                if (workbook.GetSheetIndex(name) == -1)
                {
                    return name;
                }
            }
        }
        /// <summary>
        /// Create an HSSFSheet for this HSSFWorkbook, Adds it to the sheets and
        /// returns the high level representation. Use this to Create new sheets.
        /// </summary>
        /// <param name="sheetname">sheetname to set for the sheet.</param>
        /// <returns>HSSFSheet representing the new sheet.</returns>
        public ISheet CreateSheet(String sheetname)
        {
            if (sheetname == null)
            {
                throw new ArgumentException("sheetName must not be null");
            }

            if (workbook.ContainsSheetName(sheetname, _sheets.Count))
                throw new ArgumentException("The workbook already contains a sheet of this name");

            HSSFSheet sheet = new HSSFSheet(this);

            workbook.SetSheetName(_sheets.Count, sheetname);
            _sheets.Add(sheet);

            bool isOnlySheet = _sheets.Count == 1;
            sheet.IsSelected = isOnlySheet;
            sheet.IsActive = isOnlySheet;
            return sheet;
        }

        /// <summary>
        /// Get the number of spreadsheets in the workbook (this will be three after serialization)
        /// </summary>
        /// <value>The number of sheets.</value>
        public int NumberOfSheets
        {
            get { return _sheets.Count; }
        }

        /// <summary>
        /// Gets the sheets.
        /// </summary>
        /// <returns></returns>
        private List<HSSFSheet> GetSheets()
        {
            return _sheets;
        }
        ///<summary>
        /// Get the HSSFSheet object at the given index.
        ///</summary>
        /// <param name="index">index of the sheet number (0-based)</param>
        /// <returns>HSSFSheet at the provided index</returns>
        public NPOI.SS.UserModel.ISheet GetSheetAt(int index)
        {
            return (HSSFSheet)_sheets[index];
        }

        /// <summary>
        /// Get sheet with the given name (case insensitive match)
        /// </summary>
        /// <param name="name">name of the sheet</param>
        /// <returns>HSSFSheet with the name provided or null if it does not exist</returns>
        public NPOI.SS.UserModel.ISheet GetSheet(String name)
        {
            HSSFSheet retval = null;

            for (int k = 0; k < _sheets.Count; k++)
            {
                String sheetname = workbook.GetSheetName(k);

                if (sheetname.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    retval = (HSSFSheet)_sheets[k];
                    break;
                }
            }
            return retval;
        }

        /// <summary>
        /// Removes sheet at the given index.
        /// </summary>
        /// <param name="index">index of the sheet  (0-based)</param>
        ///<remarks>
        /// Care must be taken if the Removed sheet Is the currently active or only selected sheet in
        /// the workbook. There are a few situations when Excel must have a selection and/or active
        /// sheet. (For example when printing - see Bug 40414).
        /// This method makes sure that if the Removed sheet was active, another sheet will become
        /// active in its place.  Furthermore, if the Removed sheet was the only selected sheet, another
        /// sheet will become selected.  The newly active/selected sheet will have the same index, or
        /// one less if the Removed sheet was the last in the workbook.
        /// </remarks>
        public void RemoveSheetAt(int index)
        {
            ValidateSheetIndex(index);
            bool wasSelected = GetSheetAt(index).IsSelected;

            _sheets.RemoveAt(index);
            workbook.RemoveSheet(index);

            // Set the remaining active/selected sheet
            int nSheets = _sheets.Count;
            if (nSheets < 1)
            {
                // nothing more to do if there are no sheets left
                return;
            }
            // the index of the closest remaining sheet to the one just deleted
            int newSheetIndex = index;
            if (newSheetIndex >= nSheets)
            {
                newSheetIndex = nSheets - 1;
            }

            if (wasSelected)
            {
                bool someOtherSheetIsStillSelected = false;
                for (int i = 0; i < nSheets; i++)
                {
                    if (GetSheetAt(i).IsSelected)
                    {
                        someOtherSheetIsStillSelected = true;
                        break;
                    }
                }
                if (!someOtherSheetIsStillSelected)
                {
                    SetSelectedTab(newSheetIndex);
                }
            }

            // adjust active sheet
            int active = ActiveSheetIndex;
            if (active == index)
            {
                // removed sheet was the active one, reset active sheet if there is still one left now
                SetActiveSheet(newSheetIndex);
            }
            else if (active > index)
            {
                // removed sheet was below the active one => active is one less now
                SetActiveSheet(active - 1);
            }
        }

        /// <summary>
        /// determine whether the Excel GUI will backup the workbook when saving.
        /// </summary>
        /// <value>the current Setting for backups.</value>
        public bool BackupFlag
        {
            get
            {
                BackupRecord backupRecord = workbook.BackupRecord;

                return (backupRecord.Backup == 0) ? false
                        : true;
            }
            set
            {
                BackupRecord backupRecord = workbook.BackupRecord;

                backupRecord.Backup = (value ? (short)1
                        : (short)0);
            }
        }

        /// <summary>
        /// Sets the repeating rows and columns for a sheet (as found in
        /// File-&gt;PageSetup-&gt;Sheet).  This Is function Is included in the workbook
        /// because it Creates/modifies name records which are stored at the
        /// workbook level.
        /// </summary>
        /// <param name="sheetIndex">0 based index to sheet.</param>
        /// <param name="startColumn">0 based start of repeating columns.</param>
        /// <param name="endColumn">0 based end of repeating columns.</param>
        /// <param name="startRow">0 based start of repeating rows.</param>
        /// <param name="endRow">0 based end of repeating rows.</param>
        /// <example>
        /// To set just repeating columns:
        /// workbook.SetRepeatingRowsAndColumns(0,0,1,-1-1);
        /// To set just repeating rows:
        /// workbook.SetRepeatingRowsAndColumns(0,-1,-1,0,4);
        /// To remove all repeating rows and columns for a sheet.
        /// workbook.SetRepeatingRowsAndColumns(0,-1,-1,-1,-1);
        /// </example>
        [Obsolete("use HSSFSheet#setRepeatingRows(CellRangeAddress) or HSSFSheet#setRepeatingColumns(CellRangeAddress)")]
        public void SetRepeatingRowsAndColumns(int sheetIndex,
                                               int startColumn, int endColumn,
                                               int startRow, int endRow)
        {
            HSSFSheet sheet = (HSSFSheet)GetSheetAt(sheetIndex);

            CellRangeAddress rows = null;
            CellRangeAddress cols = null;

            if (startRow != -1)
            {
                rows = new CellRangeAddress(startRow, endRow, -1, -1);
            }
            if (startColumn != -1)
            {
                cols = new CellRangeAddress(-1, -1, startColumn, endColumn);
            }

            sheet.RepeatingRows = (rows);
            sheet.RepeatingColumns = (cols);
        }
        internal int FindExistingBuiltinNameRecordIdx(int sheetIndex, byte builtinCode)
        {
            for (int defNameIndex = 0; defNameIndex < names.Count; defNameIndex++)
            {
                NameRecord r = workbook.GetNameRecord(defNameIndex);
                if (r == null)
                {
                    throw new InvalidOperationException("Unable to find all defined names to iterate over");
                }
                if (!r.IsBuiltInName || r.BuiltInName != builtinCode)
                {
                    continue;
                }
                if (r.SheetNumber - 1 == sheetIndex)
                {
                    return defNameIndex;
                }
            }
            return -1;
        }

        internal HSSFName CreateBuiltInName(byte builtinCode, int sheetIndex)
        {
            NameRecord nameRecord =
              workbook.CreateBuiltInName(builtinCode, sheetIndex + 1);
            HSSFName newName = new HSSFName(this, nameRecord, null);
            names.Add(newName);
            return newName;
        }


        internal HSSFName GetBuiltInName(byte builtinCode, int sheetIndex)
        {
            int index = FindExistingBuiltinNameRecordIdx(sheetIndex, builtinCode);
            if (index < 0)
            {
                return null;
            }
            else
            {
                return names[(index)];
            }
        }

        private bool IsRowColHeaderRecord(NameRecord r)
        {
            return r.OptionFlag == 0x20 && ("" + ((char)7)).Equals(r.NameText);
        }

        /// <summary>
        /// Create a new Font and Add it to the workbook's font table
        /// </summary>
        /// <returns>new font object</returns>
        public NPOI.SS.UserModel.IFont CreateFont()
        {
            FontRecord font = workbook.CreateNewFont();
            short fontindex = (short)(NumberOfFonts - 1);

            if (fontindex > 3)
            {
                fontindex++;   // THERE Is NO FOUR!!
            }
            if (fontindex == short.MaxValue)
            {
                throw new ArgumentException("Maximum number of fonts was exceeded");
            }
            // Ask getFontAt() to build it for us,
            //  so it gets properly cached
            return GetFontAt(fontindex);
        }

        //public NPOI.SS.UserModel.Font FindFont(NPOI.SS.UserModel.FontBoldWeight boldWeight, short color, short fontHeight,
        //                 String name, bool italic, bool strikeout,
        //                 NPOI.SS.UserModel.FontSuperScript typeOffset, NPOI.SS.UserModel.FontUnderlineType Underline)
        //{
        //    return this.FindFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, Underline);
        //}

        /// <summary>
        /// Finds a font that matches the one with the supplied attributes
        /// </summary>
        /// <param name="boldWeight">The bold weight.</param>
        /// <param name="color">The color.</param>
        /// <param name="fontHeight">Height of the font.</param>
        /// <param name="name">The name.</param>
        /// <param name="italic">if set to <c>true</c> [italic].</param>
        /// <param name="strikeout">if set to <c>true</c> [strikeout].</param>
        /// <param name="typeOffset">The type offset.</param>
        /// <param name="underline">The underline.</param>
        /// <returns></returns>
        public NPOI.SS.UserModel.IFont FindFont(short boldWeight, short color, short fontHeight,
                         String name, bool italic, bool strikeout,
                         FontSuperScript typeOffset, FontUnderlineType underline)
        {
            //        Console.WriteLine( boldWeight + ", " + color + ", " + fontHeight + ", " + name + ", " + italic + ", " + strikeout + ", " + typeOffset + ", " + Underline );
            for (short i = 0; i <= this.NumberOfFonts; i++)
            {
                // Remember - there is no 4!
                if (i == 4)
                    continue;

                NPOI.SS.UserModel.IFont hssfFont = GetFontAt(i);
                //            Console.WriteLine( hssfFont.GetBoldweight() + ", " + hssfFont.GetColor() + ", " + hssfFont.FontHeight + ", " + hssfFont.FontName + ", " + hssfFont.GetItalic() + ", " + hssfFont.GetStrikeout() + ", " + hssfFont.GetTypeOffset() + ", " + hssfFont.Underline );
                if (hssfFont.Boldweight == boldWeight
                        && hssfFont.Color == color
                        && hssfFont.FontHeight == fontHeight
                        && hssfFont.FontName.Equals(name)
                        && hssfFont.IsItalic == italic
                        && hssfFont.IsStrikeout == strikeout
                        && hssfFont.TypeOffset == typeOffset
                        && hssfFont.Underline == underline)
                {
                    //                Console.WriteLine( "Found font" );
                    return hssfFont;
                }
            }

            //        Console.WriteLine( "No font found" );
            return null;
        }

        /// <summary>
        /// Get the number of fonts in the font table
        /// </summary>
        /// <value>The number of fonts.</value>
        public short NumberOfFonts
        {
            get
            {
                return (short)workbook.NumberOfFontRecords;
            }
        }
        public bool IsHidden
        {
            get
            {
                return workbook.WindowOne.Hidden;
            }
            set
            {
                workbook.WindowOne.Hidden = value;
            }
        }


        /// <summary>
        /// Get the font at the given index number
        /// </summary>
        /// <param name="idx">The index number</param>
        /// <returns>HSSFFont at the index</returns>
        public NPOI.SS.UserModel.IFont GetFontAt(short idx)
        {
            if (fonts == null) fonts = new Hashtable();

            // So we don't confuse users, give them back
            //  the same object every time, but create
            //  them lazily

            if (fonts.ContainsKey(idx))
            {
                return (HSSFFont)fonts[idx];
            }

            FontRecord font = workbook.GetFontRecordAt(idx);
            HSSFFont retval = new HSSFFont(idx, font);
            fonts[idx] = retval;

            return retval;
        }

        /// <summary>
        /// Reset the fonts cache, causing all new calls
        /// to getFontAt() to create new objects.
        /// Should only be called after deleting fonts,
        /// and that's not something you should normally do
        /// </summary>
        public void ResetFontCache()
        {
            fonts = new Hashtable();
        }
        /// <summary>
        /// Create a new Cell style and Add it to the workbook's style table
        /// </summary>
        /// <returns>the new Cell Style object</returns>
        public NPOI.SS.UserModel.ICellStyle CreateCellStyle()
        {
            if (workbook.NumExFormats == MAX_STYLES)
            {
                throw new InvalidOperationException("The maximum number of cell styles was exceeded. " +
                        "You can define up to 4000 styles in a .xls workbook");
            }
            ExtendedFormatRecord xfr = workbook.CreateCellXF();
            short index = (short)(NumCellStyles - 1);
            HSSFCellStyle style = new HSSFCellStyle(index, xfr, this);

            return style;
        }

        /// <summary>
        /// Get the number of styles the workbook Contains
        /// </summary>
        /// <value>count of cell styles</value>
        public short NumCellStyles
        {
            get
            {
                return (short)workbook.NumExFormats;
            }

        }

        /// <summary>
        /// Get the cell style object at the given index
        /// </summary>
        /// <param name="idx">index within the Set of styles</param>
        /// <returns>HSSFCellStyle object at the index</returns>
        public NPOI.SS.UserModel.ICellStyle GetCellStyleAt(short idx)
        {
            ExtendedFormatRecord xfr = workbook.GetExFormatAt(idx);
            HSSFCellStyle style = new HSSFCellStyle(idx, xfr, this);

            return style;
        }

        /**
         * Closes the underlying {@link NPOIFSFileSystem} from which
         *  the Workbook was read, if any. Has no effect on Workbooks
         *  opened from an InputStream, or newly created ones.
         */
        public void Close()
        {
            if (directory != null)
            {
                if (directory.NFileSystem != null)
                {
                    directory.NFileSystem.Close();
                    directory = null;
                }
            }
        }

        /// <summary>
        /// Write out this workbook to an Outputstream.  Constructs
        /// a new POI POIFSFileSystem, passes in the workbook binary representation  and
        /// Writes it out.
        /// </summary>
        /// <param name="stream">the java OutputStream you wish to Write the XLS to</param>
        public override void Write(Stream stream)
        {
            byte[] bytes = GetBytes();
            POIFSFileSystem fs = new POIFSFileSystem();

            if (this.DocumentSummaryInformation == null)
            {
                this.DocumentSummaryInformation = HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
            }
            NPOI.HPSF.CustomProperties cp = this.DocumentSummaryInformation.CustomProperties;
            if(cp==null)
            {
                cp= new NPOI.HPSF.CustomProperties();
            }
            cp.Put("Generator", "NPOI");
            cp.Put("Generator Version", Assembly.GetExecutingAssembly().GetName().Version.ToString(3));
            this.DocumentSummaryInformation.CustomProperties = cp;
            if (this.SummaryInformation == null)
            {
                this.SummaryInformation = HPSF.PropertySetFactory.CreateSummaryInformation();
            }            
            this.SummaryInformation.ApplicationName = "NPOI";

            // For tracking what we've written out, used if we're
            //  going to be preserving nodes
            List<string> excepts = new List<string>(1);

            using (MemoryStream newMemoryStream = new MemoryStream(bytes))
            {
                // Write out the Workbook stream
                fs.CreateDocument(newMemoryStream, "Workbook");

                // Write out our HPFS properties, if we have them
                WriteProperties(fs, excepts);

                if (preserveNodes)
                {
                    // Don't Write out the old Workbook, we'll be doing our new one
                    // If the file had an "incorrect" name for the workbook stream,
                    // don't write the old one as we'll use the correct name shortly
                    excepts.AddRange(WORKBOOK_DIR_ENTRY_NAMES);

                    // Copy over all the other nodes to our new poifs
                    EntryUtils.CopyNodes(
                            new FilteringDirectoryNode(this.directory, excepts)
                            , new FilteringDirectoryNode(fs.Root, excepts)
                    );
                    // YK: preserve StorageClsid, it is important for embedded workbooks,
                    // see Bugzilla 47920
                    fs.Root.StorageClsid = (this.directory.StorageClsid);
                }
                fs.WriteFileSystem(stream);

            }
            
            bytes = null;
        }

        /// <summary>
        /// Get the bytes of just the HSSF portions of the XLS file.
        /// Use this to construct a POI POIFSFileSystem yourself.
        /// </summary>
        /// <returns>byte[] array containing the binary representation of this workbook and all contained
        /// sheets, rows, cells, etc.</returns>
        public byte[] GetBytes()
        {
            //if (log.Check(POILogger.DEBUG))
            //{
            //    log.Log(DEBUG, "HSSFWorkbook.GetBytes()");
            //}

            List<HSSFSheet> sheets = GetSheets();
            int nSheets = sheets.Count;

            // before Getting the workbook size we must tell the sheets that
            // serialization Is about to occur.
            workbook.PreSerialize();
            for (int i = 0; i < nSheets; i++)
            {
                sheets[i].Sheet.Preserialize();
                sheets[i].PreSerialize();
            }

            int totalsize = workbook.Size;

            // pre-calculate all the sheet sizes and set BOF indexes
            SheetRecordCollector[] srCollectors = new SheetRecordCollector[nSheets];
            for (int k = 0; k < nSheets; k++)
            {
                workbook.SetSheetBof(k, totalsize);
                using (SheetRecordCollector src = new SheetRecordCollector())
                {
                    sheets[k].Sheet.VisitContainedRecords(src, totalsize);

                    totalsize += src.TotalSize;
                    srCollectors[k] = src;
                }
            }


            byte[] retval = new byte[totalsize];
            int pos = workbook.Serialize(0, retval);

            for (int k = 0; k < nSheets; k++)
            {
                SheetRecordCollector src = srCollectors[k];
                int serializedSize = src.Serialize(pos, retval);
                if (serializedSize != src.TotalSize)
                {
                    // Wrong offset values have been passed in the call to SetSheetBof() above.
                    // For books with more than one sheet, this discrepancy would cause excel 
                    // to report errors and loose data while Reading the workbook
                    throw new InvalidOperationException("Actual serialized sheet size (" + serializedSize
                            + ") differs from pre-calculated size (" + src.TotalSize
                            + ") for sheet (" + k + ")");
                    // TODO - Add similar sanity Check to Ensure that Sheet.SerializeIndexRecord() does not Write mis-aligned offsets either
                }
                pos += serializedSize;
                src.Dispose();
            }
            
            return retval;
        }

        [Obsolete("Do not call this method from your applications. Use the methods available in the HSSFRow to Add string HSSFCells")]
        public int AddSSTString(String str)
        {
            return workbook.AddSSTString(new UnicodeString(str));
        }

        [Obsolete("Do not call this method from your applications. Use the methods available in the HSSFRow to Get string HSSFCells")]
        public String GetSSTString(int index)
        {
            return workbook.GetSSTString(index).String;
        }
        /**
 * The locator of user-defined functions.
 * By default includes functions from the Excel Analysis Toolpack
 */
        [NonSerialized]
        private UDFFinder _udfFinder = new IndexedUDFFinder(UDFFinder.DEFAULT);

        /**
 * Register a new toolpack in this workbook.
 *
 * @param toopack the toolpack to register
 */
        public void AddToolPack(UDFFinder toopack)
        {
            AggregatingUDFFinder udfs = (AggregatingUDFFinder)_udfFinder;
            udfs.Add(toopack);
        }
        /*package*/
        internal UDFFinder GetUDFFinder()
        {
            return _udfFinder;
        }

        /// <summary>
        /// Gets the workbook.
        /// </summary>
        /// <value>The workbook.</value>
        public InternalWorkbook Workbook
        {
            get { return workbook; }
        }

        /// <summary>
        /// Gets the total number of named ranges in the workboko
        /// </summary>
        /// <value>The number of named ranges</value>
        public int NumberOfNames
        {
            get
            {
                int result = names.Count;
                return result;
            }
        }
        public NPOI.SS.UserModel.IName GetName(String name)
        {
            int nameIndex = GetNameIndex(name);
            if (nameIndex < 0)
            {
                return null;
            }
            return (HSSFName)names[nameIndex];
        }

        /// <summary>
        /// Gets the Named range
        /// </summary>
        /// <param name="nameIndex">position of the named range</param>
        /// <returns>named range high level</returns>
        public NPOI.SS.UserModel.IName GetNameAt(int nameIndex)
        {
            int nNames = names.Count;
            if (nNames < 1)
            {
                throw new InvalidOperationException("There are no defined names in this workbook");
            }
            if (nameIndex < 0 || nameIndex > nNames)
            {
                throw new ArgumentOutOfRangeException("Specified name index " + nameIndex
                        + " is outside the allowable range (0.." + (nNames - 1) + ").");
            }
            HSSFName result = names[nameIndex];

            return result;
        }

        /// <summary>
        /// Gets the named range name
        /// </summary>
        /// <param name="index">the named range index (0 based)</param>
        /// <returns>named range name</returns>
        public String GetNameName(int index)
        {
            String result = GetNameAt(index).NameName;

            return result;
        }
        public NameRecord GetNameRecord(int nameIndex)
        {
            return Workbook.GetNameRecord(nameIndex);
        }
        /// <summary>
        /// TODO - make this less cryptic / move elsewhere
        /// </summary>
        /// <param name="reFindex">Index to REF entry in EXTERNSHEET record in the Link Table</param>
        /// <param name="definedNameIndex">zero-based to DEFINEDNAME or EXTERNALNAME record</param>
        /// <returns>the string representation of the defined or external name</returns>
        public String ResolveNameXText(int reFindex, int definedNameIndex)
        {
            return workbook.ResolveNameXText(reFindex, definedNameIndex);
        }


        /// <summary>
        /// Sets the printarea for the sheet provided
        /// i.e. Reference = $A$1:$B$2
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index (0 Represents the first sheet to keep consistent with java)</param>
        /// <param name="reference">Valid name Reference for the Print Area</param>
        public void SetPrintArea(int sheetIndex, String reference)
        {
            NameRecord name = workbook.GetSpecificBuiltinRecord(NameRecord.BUILTIN_PRINT_AREA, sheetIndex + 1);


            if (name == null)
                name = workbook.CreateBuiltInName(NameRecord.BUILTIN_PRINT_AREA, sheetIndex + 1);
            //Adding one here because 0 indicates a global named region; doesnt make sense for print areas
            String[] parts = reference.Split(new char[]{','});
            StringBuilder sb = new StringBuilder(32);
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0)
                {
                    sb.Append(",");
                }
                SheetNameFormatter.AppendFormat(sb, GetSheetName(sheetIndex));
                sb.Append("!");
                sb.Append(parts[i]);
            }
            name.NameDefinition =(HSSFFormulaParser.Parse(sb.ToString(), this, FormulaType.NamedRange, sheetIndex));
        }

        /// <summary>
        /// Sets the print area.
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index (0 = First Sheet)</param>
        /// <param name="startColumn">Column to begin printarea</param>
        /// <param name="endColumn">Column to end the printarea</param>
        /// <param name="startRow">Row to begin the printarea</param>
        /// <param name="endRow">Row to end the printarea</param>
        public void SetPrintArea(int sheetIndex, int startColumn, int endColumn,
                                  int startRow, int endRow)
        {

            //using absolute references because they don't Get copied and pasted anyway
            CellReference cell = new CellReference(startRow, startColumn, true, true);
            String reference = cell.FormatAsString();

            cell = new CellReference(endRow, endColumn, true, true);
            reference = reference + ":" + cell.FormatAsString();

            SetPrintArea(sheetIndex, reference);
        }


        /// <summary>
        /// Retrieves the reference for the printarea of the specified sheet, the sheet name Is Appended to the reference even if it was not specified.
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index (0 Represents the first sheet to keep consistent with java)</param>
        /// <returns>String Null if no print area has been defined</returns>
        public String GetPrintArea(int sheetIndex)
        {
            NameRecord name = workbook.GetSpecificBuiltinRecord(NameRecord.BUILTIN_PRINT_AREA, sheetIndex + 1);
            if (name == null) return null;
            //Adding one here because 0 indicates a global named region; doesnt make sense for print areas
            return HSSFFormulaParser.ToFormulaString(this, name.NameDefinition);
        }

        /// <summary>
        /// Delete the printarea for the sheet specified
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index (0 = First Sheet)</param>
        public void RemovePrintArea(int sheetIndex)
        {
            Workbook.RemoveBuiltinRecord(NameRecord.BUILTIN_PRINT_AREA, sheetIndex + 1);
        }

        /// <summary>
        /// Creates a new named range and Add it to the model
        /// </summary>
        /// <returns>named range high level</returns>
        public NPOI.SS.UserModel.IName CreateName()
        {
            NameRecord nameRecord = workbook.CreateName();

            HSSFName newName = new HSSFName(this, nameRecord);

            names.Add(newName);

            return newName;
        }

        /// <summary>
        /// Gets the named range index by his name
        /// Note:
        /// Excel named ranges are case-insensitive and
        /// this method performs a case-insensitive search.
        /// </summary>
        /// <param name="name">named range name</param>
        /// <returns>named range index</returns>
        public int GetNameIndex(String name)
        {
            int retval = -1;

            for (int k = 0; k < names.Count; k++)
            {
                String nameName = GetNameName(k);

                if (nameName.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    retval = k;
                    break;
                }
            }
            return retval;
        }

        //
        /// <summary>
        /// As GetNameIndex(String) is not necessarily unique 
        /// (name + sheet index is unique), this method is more accurate.
        /// </summary>
        /// <param name="name">the name whose index in the list of names of this workbook should be looked up.</param>
        /// <returns>an index value >= 0 if the name was found; -1, if the name was not found</returns>
        public int GetNameIndex(HSSFName name)
        {
            for (int k = 0; k < names.Count; k++)
            {
                if (name == names[(k)])
                {
                    return k;
                }
            }
            return -1;
        }

        /// <summary>
        /// Remove the named range by his index
        /// </summary>
        /// <param name="index">The named range index (0 based)</param>
        public void RemoveName(int index)
        {
            names.RemoveAt(index);
            workbook.RemoveName(index);
        }

        /// <summary>
        /// Creates the instance of HSSFDataFormat for this workbook.
        /// </summary>
        /// <returns>the HSSFDataFormat object</returns>
        public NPOI.SS.UserModel.IDataFormat CreateDataFormat()
        {
            if (formatter == null)
                formatter = new HSSFDataFormat(workbook);
            return formatter;
        }

        /// <summary>
        /// Remove the named range by his name
        /// </summary>
        /// <param name="name">named range name</param>
        public void RemoveName(String name)
        {
            int index = GetNameIndex(name);

            RemoveName(index);

        }

        //
        /// <summary>
        ///  As #removeName(String) is not necessarily unique (name + sheet index is unique), 
        ///  this method is more accurate.
        /// </summary>
        /// <param name="name">the name to remove.</param>
        public void RemoveName(HSSFName name)
        {
            int index = GetNameIndex(name);
            RemoveName(index);
        }
        public HSSFPalette GetCustomPalette()
        {
            return new HSSFPalette(workbook.CustomPalette);
        }

        // /** Test only. Do not use */
        //public void InsertChartRecord()
        //{
        //    int loc = workbook.FindFirstRecordLocBySid(SSTRecord.sid);
        //    byte[] data = {
        //       (byte)0x0F, (byte)0x00, (byte)0x00, (byte)0xF0, (byte)0x52,
        //       (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
        //       (byte)0x06, (byte)0xF0, (byte)0x18, (byte)0x00, (byte)0x00,
        //       (byte)0x00, (byte)0x01, (byte)0x08, (byte)0x00, (byte)0x00,
        //       (byte)0x02, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x02,
        //       (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x01, (byte)0x00,
        //       (byte)0x00, (byte)0x00, (byte)0x01, (byte)0x00, (byte)0x00,
        //       (byte)0x00, (byte)0x03, (byte)0x00, (byte)0x00, (byte)0x00,
        //       (byte)0x33, (byte)0x00, (byte)0x0B, (byte)0xF0, (byte)0x12,
        //       (byte)0x00, (byte)0x00, (byte)0x00, (byte)0xBF, (byte)0x00,
        //       (byte)0x08, (byte)0x00, (byte)0x08, (byte)0x00, (byte)0x81,
        //       (byte)0x01, (byte)0x09, (byte)0x00, (byte)0x00, (byte)0x08,
        //       (byte)0xC0, (byte)0x01, (byte)0x40, (byte)0x00, (byte)0x00,
        //       (byte)0x08, (byte)0x40, (byte)0x00, (byte)0x1E, (byte)0xF1,
        //       (byte)0x10, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x0D,
        //       (byte)0x00, (byte)0x00, (byte)0x08, (byte)0x0C, (byte)0x00,
        //       (byte)0x00, (byte)0x08, (byte)0x17, (byte)0x00, (byte)0x00,
        //       (byte)0x08, (byte)0xF7, (byte)0x00, (byte)0x00, (byte)0x10,
        //    };
        //    UnknownRecord r = new UnknownRecord((short)0x00EB, data);
        //    workbook.Records.Insert(loc, r);
        //}

        /// <summary>
        /// Spits out a list of all the drawing records in the workbook.
        /// </summary>
        /// <param name="fat">if set to <c>true</c> [fat].</param>
        public void DumpDrawingGroupRecords(bool fat)
        {
            DrawingGroupRecord r = (DrawingGroupRecord)workbook.FindFirstRecordBySid(DrawingGroupRecord.sid);
            r.Decode();
            IList escherRecords = r.EscherRecords;

            for (IEnumerator iterator = escherRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord escherRecord = (EscherRecord)iterator.Current;
                if (fat)
                    Console.WriteLine(escherRecord.ToString());
                else
                    escherRecord.Display(0);
            }
        }
        internal void InitDrawings()
        {
            DrawingManager2 mgr = workbook.FindDrawingGroup();
            if (mgr != null)
            {
                for (int i = 0; i < NumberOfSheets; i++)
                {
                    IDrawing tmp = GetSheetAt(i).DrawingPatriarch;
                }
            }
            else
            {
                workbook.CreateDrawingGroup();
            }
        }
        /// <summary>
        /// Adds a picture to the workbook.
        /// </summary>
        /// <param name="pictureData">The bytes of the picture</param>
        /// <param name="format">The format of the picture.  One of 
        /// PictureType.</param>
        /// <returns>the index to this picture (1 based).</returns>
        public int AddPicture(byte[] pictureData, NPOI.SS.UserModel.PictureType format)
        {
            InitDrawings();

            byte[] uid;
            using (MD5 hasher = MD5.Create())
            {
                uid = hasher.ComputeHash(pictureData);
            }
            EscherBlipRecord blipRecord;
            int blipSize;
            short escherTag;
            switch (format) {
                case PictureType.WMF:
                    // remove first 22 bytes if file starts with magic bytes D7-CD-C6-9A
                    // see also http://de.wikipedia.org/wiki/Windows_Metafile#Hinweise_zur_WMF-Spezifikation
                    if (LittleEndian.GetInt(pictureData) == unchecked((int)0x9AC6CDD7)) {
                        byte[] picDataNoHeader = new byte[pictureData.Length-22];
                        Array.Copy(pictureData, 22, picDataNoHeader, 0, pictureData.Length-22);
                        pictureData = picDataNoHeader;
                    }
                    EscherMetafileBlip blipRecordMeta = new EscherMetafileBlip();
                    blipRecord = blipRecordMeta;
                    blipRecordMeta.UID=(/*setter*/uid);
                    blipRecordMeta.SetPictureData(pictureData);
                    // taken from libre office export, it won't open, if this is left to 0
                    blipRecordMeta.Filter=(/*setter*/unchecked((byte)-2));
                    blipSize = blipRecordMeta.CompressedSize + 58;
                    escherTag = 0;
                    break;
                case PictureType.EMF:
                    blipRecordMeta = new EscherMetafileBlip();
                    blipRecord = blipRecordMeta;
                    blipRecordMeta.UID=(/*setter*/uid);
                    blipRecordMeta.SetPictureData(pictureData);
                    // taken from libre office export, it won't open, if this is left to 0
                    blipRecordMeta.Filter=(/*setter*/unchecked((byte)-2));
                    blipSize = blipRecordMeta.CompressedSize + 58;
                    escherTag = 0;
                    break;
                default:
                    EscherBitmapBlip blipRecordBitmap = new EscherBitmapBlip();
                    blipRecord = blipRecordBitmap;
                    blipRecordBitmap.UID=(/*setter*/ uid );
                    blipRecordBitmap.Marker=(/*setter*/ (byte) 0xFF );
                    blipRecordBitmap.PictureData=(pictureData);
                    blipSize = pictureData.Length + 25;
                    escherTag = (short) 0xFF;
    	            break;
            }

            blipRecord.RecordId = (short)(EscherBitmapBlip.RECORD_ID_START + format);
            
            switch (format)
            {
                case NPOI.SS.UserModel.PictureType.EMF:
                    blipRecord.Options = HSSFPictureData.MSOBI_EMF;
                    break;
                case NPOI.SS.UserModel.PictureType.WMF:
                    blipRecord.Options = HSSFPictureData.MSOBI_WMF;
                    break;
                case NPOI.SS.UserModel.PictureType.PICT:
                    blipRecord.Options = HSSFPictureData.MSOBI_PICT;
                    break;
                case NPOI.SS.UserModel.PictureType.PNG:
                    blipRecord.Options = HSSFPictureData.MSOBI_PNG;
                    break;
                case NPOI.SS.UserModel.PictureType.JPEG:
                    blipRecord.Options = HSSFPictureData.MSOBI_JPEG;
                    break;
                case NPOI.SS.UserModel.PictureType.DIB:
                    blipRecord.Options = HSSFPictureData.MSOBI_DIB;
                    break;
            }

            EscherBSERecord r = new EscherBSERecord();
            r.RecordId = EscherBSERecord.RECORD_ID;
            r.Options = (short)(0x0002 | ((int)format << 4));
            r.BlipTypeMacOS = (byte)format;
            r.BlipTypeWin32 = (byte)format;
            r.UID = uid;
            r.Tag = escherTag;
            r.Size = blipSize;
            r.Ref = 0;
            r.Offset = 0;
            r.BlipRecord = blipRecord;

            return workbook.AddBSERecord(r);
        }

        /// <summary>
        /// Gets all pictures from the Workbook.
        /// </summary>
        /// <returns>the list of pictures (a list of HSSFPictureData objects.)</returns>
        public IList GetAllPictures()
        {
            // The drawing Group record always exists at the top level, so we won't need to do this recursively.
            List<HSSFPictureData> pictures = new List<HSSFPictureData>();
            IEnumerator recordIter = workbook.Records.GetEnumerator();
            while (recordIter.MoveNext())
            {
                Object obj = recordIter.Current;
                if (obj is AbstractEscherHolderRecord)
                {
                    ((AbstractEscherHolderRecord)obj).Decode();
                    IList escherRecords = ((AbstractEscherHolderRecord)obj).EscherRecords;
                    SearchForPictures(escherRecords, pictures);
                }
            }
            return pictures;
        }
        //public HSSFAutoFilter CreateAutoFilter(String formula)
        //{
        //    if (string.IsNullOrEmpty(formula))
        //        return null;

        //    HSSFAutoFilter autofilter = new HSSFAutoFilter(formula, this);
        //    return autofilter;
        //}
//        public HSSFAutoFilter CreateCustomFilter(string formula,)

        /// <summary>
        /// Performs a recursive search for pictures in the given list of escher records.
        /// </summary>
        /// <param name="escherRecords">the escher records.</param>
        /// <param name="pictures">the list to populate with the pictures.</param>
        private void SearchForPictures(IList escherRecords, List<HSSFPictureData> pictures)
        {
            IEnumerator recordIter = escherRecords.GetEnumerator();
            while (recordIter.MoveNext())
            {
                Object obj = recordIter.Current;
                if (obj is EscherRecord)
                {
                    EscherRecord escherRecord = (EscherRecord)obj;

                    if (escherRecord is EscherBSERecord)
                    {
                        EscherBlipRecord blip = ((EscherBSERecord)escherRecord).BlipRecord;
                        if (blip != null)
                        {
                            // TODO: Some kind of structure.
                            pictures.Add(new HSSFPictureData(blip));
                        }
                    }

                    // Recursive call.
                    SearchForPictures(escherRecord.ChildRecords, pictures);
                }
            }
        }
        protected static Dictionary<String, ClassID> GetOleMap()
        {
            Dictionary<String, ClassID> olemap = new Dictionary<String, ClassID>();
            olemap.Add("PowerPoint Document", ClassID.PPT_SHOW);
            foreach (String str in WORKBOOK_DIR_ENTRY_NAMES)
            {
                olemap.Add(str, ClassID.XLS_WORKBOOK);
            }
            // ... to be continued
            return olemap;
        }

        public int AddOlePackage(POIFSFileSystem poiData, String label, String fileName, String command)
        {
            DirectoryNode root = poiData.Root;
            Dictionary<String, ClassID> olemap = GetOleMap();
            foreach (KeyValuePair<String, ClassID> entry in olemap)
            {
                if (root.HasEntry(entry.Key))
                {
                    root.StorageClsid = (/*setter*/entry.Value);
                    break;
                }
            }

            MemoryStream bos = new MemoryStream();
            poiData.WriteFileSystem(bos);
            return AddOlePackage(bos.ToArray(), label, fileName, command);
        }

        public int AddOlePackage(byte[] oleData, String label, String fileName, String command)
        {
            // check if we were Created by POIFS otherwise create a new dummy POIFS for storing the package data
            if (directory == null)
            {
                directory = new POIFSFileSystem().Root;
                preserveNodes = true;
            }

            // Get free MBD-Node
            int storageId = 0;
            DirectoryEntry oleDir = null;
            do
            {
                String storageStr = "MBD" + HexDump.ToHex(++storageId);
                if (!directory.HasEntry(storageStr))
                {
                    oleDir = directory.CreateDirectory(storageStr);
                    oleDir.StorageClsid = (/*setter*/ClassID.OLE10_PACKAGE);
                }
            } while (oleDir == null);

            // the following data was taken from an example libre office document
            // beside this "\u0001Ole" record there were several other records, e.g. CompObj,
            // OlePresXXX, but it seems, that they aren't neccessary
            byte[] oleBytes = { 1, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            oleDir.CreateDocument("\u0001Ole", new MemoryStream(oleBytes));

            Ole10Native oleNative = new Ole10Native(label, fileName, command, oleData);
            MemoryStream bos = new MemoryStream();
            oleNative.WriteOut(bos);
            oleDir.CreateDocument(Ole10Native.OLE10_NATIVE, new MemoryStream(bos.ToArray()));

            return storageId;
        }

        /// <summary>
        /// Adds the LinkTable records required to allow formulas referencing
        /// the specified external workbook to be added to this one. Allows
        /// formulas such as "[MyOtherWorkbook]Sheet3!$A$5" to be added to the 
        /// file, for workbooks not already referenced.
        /// </summary>
        /// <param name="name">The name the workbook will be referenced as in formulas</param>
        /// <param name="workbook">The open workbook to fetch the link required information from</param>
        /// <returns></returns>
        public int LinkExternalWorkbook(String name, IWorkbook workbook)
        {
            return this.workbook.LinkExternalWorkbook(name, workbook);
        }

        /// <summary>
        /// Is the workbook protected with a password (not encrypted)?
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is write protected; otherwise, <c>false</c>.
        /// </value>
        public bool IsWriteProtected
        {
            get { return this.workbook.IsWriteProtected; }
        }

        /// <summary>
        /// protect a workbook with a password (not encypted, just Sets Writeprotect
        /// flags and the password.
        /// </summary>
        /// <param name="password">password to set</param>
        /// <param name="username">The username.</param>
        public void WriteProtectWorkbook(String password, String username)
        {
            this.workbook.WriteProtectWorkbook(password, username);
        }

        /// <summary>
        /// Removes the Write protect flag
        /// </summary>
        public void UnwriteProtectWorkbook()
        {
            this.workbook.UnwriteProtectWorkbook();
        }

        /// <summary>
        /// Gets all embedded OLE2 objects from the Workbook.
        /// </summary>
        /// <returns>the list of embedded objects (a list of HSSFObjectData objects.)</returns>
        public IList<HSSFObjectData> GetAllEmbeddedObjects()
        {
            List<HSSFObjectData> objects = new List<HSSFObjectData>();
            for (int i = 0; i < NumberOfSheets; i++)
            {
                GetAllEmbeddedObjects((HSSFSheet)GetSheetAt(i), objects);
            }
            return objects;
        }

        /// <summary>
        /// Gets all embedded OLE2 objects from the Workbook.
        /// </summary>
        /// <param name="sheet">the list of records to search.</param>
        /// <param name="objects">the list of embedded objects to populate.</param>
        private void GetAllEmbeddedObjects(HSSFSheet sheet, List<HSSFObjectData> objects)
        {
            HSSFPatriarch patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            if (null == patriarch)
            {
                return;
            }
            GetAllEmbeddedObjects(patriarch, objects);
        }

        /// <summary>
        /// Recursively iterates a shape container to get all embedded objects.
        /// </summary>
        /// <param name="parent">the parent.</param>
        /// <param name="objects">the list of embedded objects to populate.</param>
        private void GetAllEmbeddedObjects(HSSFShapeContainer parent, List<HSSFObjectData> objects)
        {
            foreach (HSSFShape shape in parent.Children) {
                if (shape is HSSFObjectData) {
                    objects.Add((HSSFObjectData) shape);
                } else if (shape is HSSFShapeContainer) {
                    GetAllEmbeddedObjects((HSSFShapeContainer) shape, objects);
                }
            }
        }

        /// <summary>
        /// Gets the new UID.
        /// </summary>
        /// <value>The new UID.</value>
        public byte[] NewUID
        {
            get
            {
                byte[] bytes = new byte[16];
                return bytes;
            }
        }

        /// <summary>
        /// Support foreach ISheet, e.g.
        /// HSSFWorkbook workbook = new HSSFWorkbook();
        /// foreach(ISheet sheet in workbook) ...
        /// </summary>
        /// <returns>Enumeration of all the sheets of this workbook</returns>
        public IEnumerator GetEnumerator()
        {
            return _sheets.GetEnumerator();
        }


        /// <summary>
        /// Whether the application shall perform a full recalculation when the workbook is opened.
        /// 
        /// Typically you want to force formula recalculation when you modify cell formulas or values
        /// of a workbook previously created by Excel. When set to true, this flag will tell Excel
        /// that it needs to recalculate all formulas in the workbook the next time the file is opened.
        /// 
        /// Note, that recalculation updates cached formula results and, thus, modifies the workbook.
        /// Depending on the version, Excel may prompt you with "Do you want to save the changes in <em>filename</em>?"
        /// on close.
        /// 
        /// Value is true if the application will perform a full recalculation of
        /// workbook values when the workbook is opened.
        /// 
        /// since 3.8
        /// </summary>
        public bool ForceFormulaRecalculation
        {
            set
            {
                InternalWorkbook iwb = Workbook;
                RecalcIdRecord recalc = iwb.RecalcId;
                recalc.EngineId=(0);
            }
            get
            {
                InternalWorkbook iwb = Workbook;
                RecalcIdRecord recalc = (RecalcIdRecord)iwb.FindFirstRecordBySid(RecalcIdRecord.sid);
                return recalc != null && recalc.EngineId != 0;
            }
        }

        /**
         * Changes an external referenced file to another file.
         * A formular in Excel which refers a cell in another file is saved in two parts: 
         * The referenced file is stored in an reference table. the row/cell information is saved separate.
         * This method invokation will only change the reference in the lookup-table itself.
         * @param oldUrl The old URL to search for and which is to be replaced
         * @param newUrl The URL replacement
         * @return true if the oldUrl was found and replaced with newUrl. Otherwise false
         */
        public bool ChangeExternalReference(String oldUrl, String newUrl)
        {
            return workbook.ChangeExternalReference(oldUrl, newUrl);
        }

        public DirectoryNode RootDirectory
        {
            get
            {
                return directory;
            }
        }

        public int IndexOf(ISheet item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, ISheet item)
        {
            this._sheets.Insert(index, (HSSFSheet)item);
        }

        public void RemoveAt(int index)
        {
            this._sheets.RemoveAt(index);
        }

        public ISheet this[int index]
        {
            get
            {
                return this.GetSheetAt(index);
            }
            set
            {
                if (this._sheets[index] != null)
                {
                    this._sheets[index] = (HSSFSheet)value;
                }
                else
                {
                    this._sheets.Insert(index, (HSSFSheet)value);
                }
            }
        }

        public void Add(ISheet item)
        {
            this._sheets.Add((HSSFSheet)item);
        }

        public void Clear()
        {
            this._sheets.Clear();
        }

        public bool Contains(ISheet item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(ISheet[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return this.NumberOfSheets; }
        }

        public bool IsReadOnly
        {
            get { throw new NotImplementedException(); }
        }

        public bool Remove(ISheet item)
        {
            return this._sheets.Remove((HSSFSheet)item);
        }
    }
}

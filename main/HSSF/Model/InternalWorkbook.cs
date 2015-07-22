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
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Util;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Udf;
    using NPOI.SS.UserModel;
    using System.Security;


    /**
     * Low level model implementation of a Workbook.  Provides creational methods
     * for Settings and objects contained in the workbook object.
     * 
     * This file Contains the low level binary records starting at the workbook's BOF and
     * ending with the workbook's EOF.  Use HSSFWorkbook for a high level representation.
     * 
     * The structures of the highlevel API use references to this to perform most of their
     * operations.  Its probably Unwise to use these low level structures directly Unless you
     * really know what you're doing.  I recommend you Read the Microsoft Excel 97 Developer's
     * Kit (Microsoft Press) and the documentation at http://sc.openoffice.org/excelfileformat.pdf
     * before even attempting to use this.
     *
     *
     * @author  Luc Girardin (luc dot girardin at macrofocus dot com)
     * @author  Sergei Kozello (sergeikozello at mail.ru)
     * @author  Shawn Laubach (slaubach at apache dot org) (Data Formats)
     * @author  Andrew C. Oliver (acoliver at apache dot org)
     * @author  Brian Sanders (bsanders at risklabs dot com) - custom palette
     * @author  Dan Sherman (dsherman at Isisph.com)
     * @author  Glen Stampoultzis (glens at apache.org)
     * @see org.apache.poi.hssf.usermodel.HSSFWorkbook
     * @version 1.0-pre
     */
    [Serializable]
    public class InternalWorkbook
    {
        /**
         * Excel silently truncates long sheet names to 31 chars.
         * This constant is used to ensure uniqueness in the first 31 chars
         */
        private const int MAX_SENSITIVE_SHEET_NAME_LEN = 31;

        //private static int DEBUG = POILogger.DEBUG;

        //    public static Workbook currentBook = null;

        /**
         * constant used to Set the "codepage" wherever "codepage" is Set in records
         * (which is duplciated in more than one record)
         */

        private const short CODEPAGE = (short)0x4b0;

        /**
         * this Contains the Worksheet record objects
         */
        [NonSerialized]
        protected WorkbookRecordList records = new WorkbookRecordList();

        /**
         * this Contains a reference to the SSTRecord so that new stings can be Added
         * to it.
         */
        [NonSerialized]
        protected SSTRecord sst = null;


        [NonSerialized]
        private LinkTable linkTable; // optionally occurs if there are  references in the document. (4.10.3)

        /**
         * holds the "boundsheet" records (aka bundlesheet) so that they can have their
         * reference to their "BOF" marker
         */
        protected List<BoundSheetRecord> boundsheets ;

        protected List<FormatRecord> formats;

        protected List<HyperlinkRecord> hyperlinks;

        protected int numxfs = 0;   // hold the number of extended format records
        protected int numfonts = 0;   // hold the number of font records
        private int maxformatid = -1;  // holds the max format id
        private bool uses1904datewindowing = false;  // whether 1904 date windowing is being used
        [NonSerialized]
        private DrawingManager2 drawingManager;
        private IList escherBSERecords ;  // EscherBSERecord
        [NonSerialized]
        private WindowOneRecord windowOne;
        [NonSerialized]
        private FileSharingRecord fileShare;
        [NonSerialized]
        private WriteAccessRecord writeAccess;
        [NonSerialized]
        private WriteProtectRecord writeProtect;

        //private static POILogger log = POILogFactory.GetLogger(typeof(Workbook));

        /**
         * Creates new Workbook with no intitialization --useless right now
         * @see #CreateWorkbook(List)
         */
        public InternalWorkbook()
        {
            records = new WorkbookRecordList();

            boundsheets = new List<BoundSheetRecord>();
            formats = new List<FormatRecord>();
            hyperlinks = new List<HyperlinkRecord>();
            numxfs = 0;
            numfonts = 0;
            maxformatid = -1;
            uses1904datewindowing = false;
            escherBSERecords = new List<EscherBSERecord>();
            commentRecords = new Dictionary<String, NameCommentRecord>();
        }

        /**
         * Read support  for low level
         * API.  Pass in an array of Record objects, A Workbook
         * object is constructed and passed back with all of its initialization Set
         * to the passed in records and references to those records held. Unlike Sheet
         * workbook does not use an offset (its assumed to be 0) since its first in a file.
         * If you need an offset then construct a new array with a 0 offset or Write your
         * own ;-p.
         *
         * @param recs an array of Record objects
         * @return Workbook object
         */
        public static InternalWorkbook CreateWorkbook(List<Record> recs)
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "Workbook (Readfile) Created with reclen=",
            //           recs.Count);
            InternalWorkbook retval = new InternalWorkbook();
            List<Record> records = new List<Record>(recs.Count / 3);
            retval.records.Records=records;

            int k;
            for (k = 0; k < recs.Count; k++)
            {
                Record rec = (Record)recs[k];

                if (rec.Sid == EOFRecord.sid)
                {
                    records.Add(rec);
                    //if (log.Check(POILogger.DEBUG))
                    //    log.Log(DEBUG, "found workbook eof record at " + k);
                    break;
                }
                switch (rec.Sid)
                {

                    case BoundSheetRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found boundsheet record at " + k);
                        retval.boundsheets.Add((BoundSheetRecord)rec);
                        retval.records.Bspos = k;
                        break;

                    case SSTRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found sst record at " + k);
                        retval.sst = (SSTRecord)rec;
                        break;

                    case FontRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found font record at " + k);
                        retval.records.Fontpos = k;
                        retval.numfonts++;
                        break;

                    case ExtendedFormatRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found XF record at " + k);
                        retval.records.Xfpos = k;
                        retval.numxfs++;
                        break;

                    case TabIdRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found tabid record at " + k);
                        retval.records.Tabpos = k;
                        break;

                    case ProtectRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found protect record at " + k);
                        retval.records.Protpos = k;
                        break;

                    case BackupRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found backup record at " + k);
                        retval.records.Backuppos = k;
                        break;
                    case ExternSheetRecord.sid:
                        throw new Exception("Extern sheet is part of LinkTable");
                    case NameRecord.sid:
                    case SupBookRecord.sid:
                        // LinkTable can start with either of these
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found SupBook record at " + k);
                        retval.linkTable = new LinkTable(recs, k, retval.records, retval.commentRecords);
                        k += retval.linkTable.RecordCount - 1;
                        continue;
                    case FormatRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found format record at " + k);
                        retval.formats.Add((FormatRecord)rec);
                        retval.maxformatid = retval.maxformatid >= ((FormatRecord)rec).IndexCode ? retval.maxformatid : ((FormatRecord)rec).IndexCode;
                        break;
                    case DateWindow1904Record.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found datewindow1904 record at " + k);
                        retval.uses1904datewindowing = ((DateWindow1904Record)rec).Windowing == 1;
                        break;
                    case PaletteRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found palette record at " + k);
                        retval.records.Palettepos = k;
                        break;
                    case WindowOneRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found WindowOneRecord at " + k);
                        retval.windowOne = (WindowOneRecord)rec;
                        break;
                    case WriteAccessRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found WriteAccess at " + k);
                        retval.writeAccess = (WriteAccessRecord)rec;
                        break;
                    case WriteProtectRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found WriteProtect at " + k);
                        retval.writeProtect = (WriteProtectRecord)rec;
                        break;
                    case FileSharingRecord.sid:
                        //if (log.Check(POILogger.DEBUG))
                        //    log.Log(DEBUG, "found FileSharing at " + k);
                        retval.fileShare = (FileSharingRecord)rec;
                        break;
                    case NameCommentRecord.sid:
                        NameCommentRecord ncr = (NameCommentRecord)rec;
                        retval.commentRecords[ncr.NameText] = ncr;
                        break;
                }
                records.Add(rec);
            }
            //What if we dont have any ranges and supbooks
            //        if (retval.records.supbookpos == 0) {
            //            retval.records.supbookpos = retval.records.bspos + 1;
            //            retval.records.namepos    = retval.records.supbookpos + 1;
            //        }

            // Look for other interesting values that
            //  follow the EOFRecord
            for (; k < recs.Count; k++)
            {
                Record rec = (Record)recs[k];
                switch (rec.Sid)
                {
                    case HyperlinkRecord.sid:
                        retval.hyperlinks.Add((HyperlinkRecord)rec);
                        break;
                }
            }

            if (retval.windowOne == null)
            {
                retval.windowOne = (WindowOneRecord)CreateWindowOne();
            }
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "exit Create workbook from existing file function");
            return retval;
        }
        
    /** gets the name comment record
     * @param nameRecord name record who's comment is required.
     * @return name comment record or <code>null</code> if there isn't one for the given name.
     */
        public NameCommentRecord GetNameCommentRecord(NameRecord nameRecord)
        {
            if (commentRecords.ContainsKey(nameRecord.NameText))
                return commentRecords[nameRecord.NameText];
            else
                return null;
        }
        /**
         * Creates an empty workbook object with three blank sheets and all the empty
         * fields.  Use this to Create a workbook from scratch.
         */
        public static InternalWorkbook CreateWorkbook()
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "creating new workbook from scratch");
            InternalWorkbook retval = new InternalWorkbook();
            List<Record> records = new List<Record>(30);
            retval.records.Records=records;
            List<FormatRecord> formats = new List<FormatRecord>(8);

            records.Add(CreateBOF());
            records.Add(new InterfaceHdrRecord(CODEPAGE));
            records.Add(CreateMMS());
            records.Add(InterfaceEndRecord.Instance);
            records.Add(CreateWriteAccess());
            records.Add(CreateCodepage());
            records.Add(CreateDSF());
            records.Add(CreateTabId());
            retval.records.Tabpos=records.Count - 1;
            records.Add(CreateFnGroupCount());
            records.Add(CreateWindowProtect());
            records.Add(CreateProtect());
            retval.records.Protpos=records.Count - 1;
            records.Add(CreatePassword());
            records.Add(CreateProtectionRev4());
            records.Add(CreatePasswordRev4());
            retval.windowOne = (WindowOneRecord)CreateWindowOne();
            records.Add(retval.windowOne);
            records.Add(CreateBackup());
            retval.records.Backuppos=records.Count - 1;
            records.Add(CreateHideObj());
            records.Add(CreateDateWindow1904());
            records.Add(CreatePrecision());
            records.Add(CreateRefreshAll());
            records.Add(CreateBookBool());
            records.Add(CreateFont());
            records.Add(CreateFont());
            records.Add(CreateFont());
            records.Add(CreateFont());
            retval.records.Fontpos=records.Count - 1;   // last font record postion
            retval.numfonts = 4;

            // Set up format records
            for (int i = 0; i <= 7; i++)
            {
                Record rec = CreateFormat(i);
                retval.maxformatid = retval.maxformatid >= ((FormatRecord)rec).IndexCode ? retval.maxformatid : ((FormatRecord)rec).IndexCode;
                formats.Add((FormatRecord)rec);
                records.Add(rec);
            }
            retval.formats = formats;

            for (int k = 0; k < 21; k++)
            {
                records.Add(CreateExtendedFormat(k));
                retval.numxfs++;
            }
            retval.records.Xfpos=records.Count - 1;
            for (int k = 0; k < 6; k++)
            {
                records.Add(CreateStyle(k));
            }
            records.Add(CreateUseSelFS());

            int nBoundSheets = 1; // now just do 1
            for (int k = 0; k < nBoundSheets; k++)
            {
                BoundSheetRecord bsr =
                        (BoundSheetRecord)CreateBoundSheet(k);

                records.Add(bsr);
                retval.boundsheets.Add(bsr);
                retval.records.Bspos=records.Count - 1;
            }
            //        retval.records.supbookpos = retval.records.bspos + 1;
            //        retval.records.namepos = retval.records.supbookpos + 2;
            records.Add(CreateCountry());
            for (int k = 0; k < nBoundSheets; k++)
            {
                retval.OrCreateLinkTable.CheckExternSheet(k);
            }
            retval.sst = new SSTRecord();
            records.Add(retval.sst);
            records.Add(CreateExtendedSST());

            records.Add(EOFRecord.instance);
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "exit Create new workbook from scratch");
            return retval;
        }


        /**Retrieves the Builtin NameRecord that matches the name and index
         * There shouldn't be too many names to make the sequential search too slow
         * @param name byte representation of the builtin name to match
         * @param sheetIndex Index to match
         * @return null if no builtin NameRecord matches
         */
        public NameRecord GetSpecificBuiltinRecord(byte name, int sheetIndex)
        {
            return OrCreateLinkTable.GetSpecificBuiltinRecord(name, sheetIndex);
        }

        public ExternalName GetExternalName(int externSheetIndex, int externNameIndex)
        {
            String nameName = linkTable.ResolveNameXText(externSheetIndex, externNameIndex, this);
            if (nameName == null)
            {
                return null;
            }
            int ix = linkTable.ResolveNameXIx(externSheetIndex, externNameIndex);
            return new ExternalName(nameName, externNameIndex, ix);
        }
        /**
         * Removes the specified Builtin NameRecord that matches the name and index
         * @param name byte representation of the builtin to match
         * @param sheetIndex zero-based sheet reference
         */
        public void RemoveBuiltinRecord(byte name, int sheetIndex)
        {
            linkTable.RemoveBuiltinRecord(name, sheetIndex);
            // TODO - do we need "this.records.Remove(...);" similar to that in this.RemoveName(int namenum) {}?
        }

        public int NumRecords
        {
            get
            {
                return records.Count;
            }
        }

        /**
         * Gets the font record at the given index in the font table.  Remember
         * "There is No Four" (someone at M$ must have gone to Rocky Horror one too
         * many times)
         *
         * @param idx the index to look at (0 or greater but NOT 4)
         * @return FontRecord located at the given index
         */

        public FontRecord GetFontRecordAt(int idx)
        {
            int index = idx;

            if (index > 4)
            {
                index -= 1;   // adjust for "There is no 4"
            }
            if (index > (numfonts - 1))
            {
                throw new IndexOutOfRangeException(
                "There are only " + numfonts
                + " font records, you asked for " + idx);
            }
            FontRecord retval =
            (FontRecord)records[(records.Fontpos - (numfonts - 1) + index)];

            return retval;
        }

        /**
         * Creates a new font record and Adds it to the "font table".  This causes the
         * boundsheets to move down one, extended formats to move down (so this function moves
         * those pointers as well)
         *
         * @return FontRecord that was just Created
         */

        public FontRecord CreateNewFont()
        {
            FontRecord rec = (FontRecord)CreateFont();

            records.Add(records.Fontpos + 1, rec);
            records.Fontpos=(records.Fontpos + 1);
            numfonts++;
            return rec;
        }

            /**
     * Check if the cloned sheet has drawings. If yes, then allocate a new drawing group ID and
     * re-generate shape IDs
     *
     * @param sheet the cloned sheet
     */
        public void CloneDrawings(InternalSheet sheet)
        {

            FindDrawingGroup();

            if (drawingManager == null)
            {
                //this workbook does not have drawings
                return;
            }

            //check if the cloned sheet has drawings
            int aggLoc = sheet.AggregateDrawingRecords(drawingManager, false);
            if (aggLoc != -1)
            {
                EscherAggregate agg = (EscherAggregate)sheet.FindFirstRecordBySid(EscherAggregate.sid);
                EscherContainerRecord escherContainer = agg.GetEscherContainer();
                if (escherContainer == null)
                {
                    return;
                }

                EscherDggRecord dgg = drawingManager.GetDgg();

                //register a new drawing group for the cloned sheet
                int dgId = drawingManager.FindNewDrawingGroupId();
                dgg.AddCluster(dgId, 0);
                dgg.DrawingsSaved = dgg.DrawingsSaved + 1;

                EscherDgRecord dg = null;
                for (IEnumerator it = escherContainer.ChildRecords.GetEnumerator(); it.MoveNext(); )
                {
                    Object er = it.Current;
                    if (er is EscherDgRecord)
                    {
                        dg = (EscherDgRecord)er;
                        //update id of the drawing in the cloned sheet
                        dg.Options = ((short)(dgId << 4));
                    }
                    else if (er is EscherContainerRecord)
                    {
                        //recursively find shape records and re-generate shapeId
                        ArrayList spRecords = new ArrayList();
                        EscherContainerRecord cp = (EscherContainerRecord)er;
                        for (IEnumerator spIt = cp.ChildRecords.GetEnumerator(); spIt.MoveNext(); )
                        {
                            EscherContainerRecord shapeContainer = (EscherContainerRecord)spIt.Current;

                            foreach (EscherRecord shapeChildRecord in shapeContainer.ChildRecords)
                            {
                                int recordId = shapeChildRecord.RecordId;
                                if (recordId == EscherSpRecord.RECORD_ID)
                                {
                                    EscherSpRecord sp = (EscherSpRecord)shapeChildRecord;
                                    int shapeId = drawingManager.AllocateShapeId((short)dgId, dg);
                                    //allocateShapeId increments the number of shapes. roll back to the previous value
                                    dg.NumShapes = (dg.NumShapes - 1);
                                    sp.ShapeId = (shapeId);
                                }
                                else if (recordId == EscherOptRecord.RECORD_ID)
                                {
                                    EscherOptRecord opt = (EscherOptRecord)shapeChildRecord;
                                    EscherSimpleProperty prop = (EscherSimpleProperty)opt.Lookup(
                                            EscherProperties.BLIP__BLIPTODISPLAY);
                                    if (prop != null)
                                    {
                                        int pictureIndex = prop.PropertyValue;
                                        // increment reference count for pictures
                                        EscherBSERecord bse = GetBSERecord(pictureIndex);
                                        bse.Ref = bse.Ref + 1;
                                    }

                                }

                            }
                        }
                    }
                }

            }
        }

        /**
         * Gets the number of font records
         *
         * @return   number of font records in the "font table"
         */

        public int NumberOfFontRecords
        {
            get
            {
                return numfonts;
            }
        }

        /**
         * Sets the BOF for a given sheet
         *
         * @param sheetnum the number of the sheet to Set the positing of the bof for
         * @param pos the actual bof position
         */

        public void SetSheetBof(int sheetIndex, int pos)
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "Setting bof for sheetnum =", sheetIndex,
            //        " at pos=",pos);
            CheckSheets(sheetIndex);
            GetBoundSheetRec(sheetIndex).PositionOfBof=pos;
        }

        /**
         * Returns the position of the backup record.
         */

        public BackupRecord BackupRecord
        {
            get
            {
                return (BackupRecord)records[records.Backuppos];
            }
        }


        /**
         * Sets the name for a given sheet.  If the boundsheet record doesn't exist and
         * its only one more than we have, go ahead and Create it.  If its > 1 more than
         * we have, except
         *
         * @param sheetnum the sheet number (0 based)
         * @param sheetname the name for the sheet
         */
        public void SetSheetName(int sheetnum, String sheetname)
        {
            CheckSheets(sheetnum);

            // YK: Mimic Excel and silently truncate sheet names longer than 31 characters
            if (sheetname.Length > 31) 
                sheetname = sheetname.Substring(0, 31);

            BoundSheetRecord sheet =boundsheets[sheetnum];
            sheet.Sheetname=sheetname;
        }

        private BoundSheetRecord GetBoundSheetRec(int sheetIndex)
        {
            return boundsheets[sheetIndex];
        }

        /**
         * Determines whether a workbook Contains the provided sheet name.
         *
         * @param name the name to test (case insensitive match)
         * @param excludeSheetIdx the sheet to exclude from the Check or -1 to include all sheets in the Check.
         * @return true if the sheet Contains the name, false otherwise.
         */
        public bool ContainsSheetName(String name, int excludeSheetIdx)
        {
            String aName = name;
            if (aName.Length > MAX_SENSITIVE_SHEET_NAME_LEN)
            {
                aName = aName.Substring(0, MAX_SENSITIVE_SHEET_NAME_LEN);
            }
            for (int i = 0; i < boundsheets.Count; i++)
            {
                BoundSheetRecord boundSheetRecord = GetBoundSheetRec(i);
                if (excludeSheetIdx == i)
                {
                    continue;
                }
                String bName = boundSheetRecord.Sheetname;
                if (bName.Length > MAX_SENSITIVE_SHEET_NAME_LEN)
                {
                    bName = bName.Substring(0, MAX_SENSITIVE_SHEET_NAME_LEN);
                }
                if (aName.Equals(bName,StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /**
         * Sets the name for a given sheet forcing the encoding. This is STILL A BAD IDEA.
         * Poi now automatically detects Unicode
         *
         *@deprecated 3-Jan-06 Simply use SetSheetNam e(int sheetnum, String sheetname)
         * @param sheetnum the sheet number (0 based)
         * @param sheetname the name for the sheet
         */
        public void SetSheetName(int sheetnum, String sheetname, short encoding)
        {
            CheckSheets(sheetnum);
            BoundSheetRecord sheet = boundsheets[sheetnum];
            sheet.Sheetname=(sheetname);
        }

        /**
         * Sets the order of appearance for a given sheet.
         *
         * @param sheetname the name of the sheet to reorder
         * @param pos the position that we want to Insert the sheet into (0 based)
         */

        public void SetSheetOrder(String sheetname, int pos)
        {
            int sheetNumber = GetSheetIndex(sheetname);
            //Remove the sheet that needs to be reordered and place it in the spot we want
            BoundSheetRecord sheet = boundsheets[sheetNumber];
            boundsheets.RemoveAt(sheetNumber);
            boundsheets.Insert(pos, sheet);

            // also adjust order of Records, calculate the position of the Boundsheets via getBspos()...
            int pos0 = records.Bspos - (boundsheets.Count - 1);
            Record removed = records[(pos0 + sheetNumber)];
            records.Remove(pos0 + sheetNumber);
            records.Add(pos0 + pos, removed);
        }

        /**
         * Gets the name for a given sheet.
         *
         * @param sheetnum the sheet number (0 based)
         * @return sheetname the name for the sheet
         */

        public String GetSheetName(int sheetIndex)
        {
            return GetBoundSheetRec(sheetIndex).Sheetname;
        }

        /**
         * Gets the hidden flag for a given sheet.
         *
         * @param sheetnum the sheet number (0 based)
         * @return True if sheet is hidden
         */

        public bool IsSheetHidden(int sheetnum)
        {
            return GetBoundSheetRec(sheetnum).IsHidden;
        }
        /**
         * Gets the hidden flag for a given sheet.
         * Note that a sheet could instead be 
         *  set to be very hidden, which is different
         *  ({@link #isSheetVeryHidden(int)})
         *
         * @param sheetnum the sheet number (0 based)
         * @return True if sheet is hidden
         */
        public bool IsSheetVeryHidden(int sheetnum)
        {
            return GetBoundSheetRec(sheetnum).IsVeryHidden;
        }
        /**
         * Hide or Unhide a sheet
         * 
         * @param sheetnum The sheet number
         * @param hidden True to mark the sheet as hidden, false otherwise
         */

        public void SetSheetHidden(int sheetnum, bool hidden)
        {
            BoundSheetRecord bsr = boundsheets[sheetnum];
            bsr.IsHidden=hidden;
        }
        /**
 * Hide or unhide a sheet.
 *  0 = not hidden
 *  1 = hidden
 *  2 = very hidden.
 * 
 * @param sheetnum The sheet number
 * @param hidden 0 for not hidden, 1 for hidden, 2 for very hidden
 */
        public void SetSheetHidden(int sheetnum, int hidden)
        {
            BoundSheetRecord bsr = GetBoundSheetRec(sheetnum);
            bool h = false;
            bool vh = false;
            if (hidden == 0)
            {
            }
            else if (hidden == 1)
            {
                h = true;
            }
            else if (hidden == 2)
            {
                vh = true;
            }
            else
            {
                throw new ArgumentException("Invalid hidden flag " + hidden + " given, must be 0, 1 or 2");
            }
            bsr.IsHidden = (h);
            bsr.IsVeryHidden = (vh);
        }
        /**
         * Get the sheet's index
         * @param name  sheet name
         * @return sheet index or -1 if it was not found.
         */

        public int GetSheetIndex(String name)
        {
            int retval = -1;

            for (int k = 0; k < boundsheets.Count; k++)
            {
                String sheet = GetSheetName(k);

                if (sheet.Equals(name,StringComparison.OrdinalIgnoreCase))
                {
                    retval = k;
                    break;
                }
            }
            return retval;
        }

        /**
         * if we're trying to Address one more sheet than we have, go ahead and Add it!  if we're
         * trying to Address >1 more than we have throw an exception!
         */

        private void CheckSheets(int sheetnum)
        {
            if ((boundsheets.Count) <= sheetnum)
            {   // if we're short one Add another..
                if ((boundsheets.Count + 1) <= sheetnum)
                {
                    throw new Exception("Sheet number out of bounds!");
                }
                BoundSheetRecord bsr = (BoundSheetRecord)CreateBoundSheet(sheetnum);

                records.Add(records.Bspos + 1, bsr);
                records.Bspos = records.Bspos + 1;
                boundsheets.Add(bsr);
                OrCreateLinkTable.CheckExternSheet(sheetnum);
                FixTabIdRecord();
            }
            //else
            //{
            //    // Ensure we have enough tab IDs
            //    // Can be a few short if new sheets were added
            //    if (records.Tabpos > 0)
            //    {
            //        TabIdRecord tir = (TabIdRecord)records[records.Tabpos];
            //        if (tir._tabids.Length < boundsheets.Count)
            //        {
            //            FixTabIdRecord();
            //        }
            //    }
            //}
        }

        public void RemoveSheet(int sheetIndex)
        {
            if (boundsheets.Count > sheetIndex)
            {
                records.Remove(records.Bspos - (boundsheets.Count - 1) + sheetIndex);
                //            records.bspos--;
                boundsheets.RemoveAt(sheetIndex);
                FixTabIdRecord();
            }

            // Within NameRecords, it's ok to have the formula
            //  part point at deleted sheets. It's also ok to
            //  have the ExternSheetNumber point at deleted
            //  sheets. 
            // However, the sheet index must be adjusted, or
            //  excel will break. (Sheet index is either 0 for
            //  global, or 1 based index to sheet)
            int sheetNum1Based = sheetIndex + 1;
            for (int i = 0; i < NumNames; i++)
            {
                NameRecord nr = GetNameRecord(i);

                if (nr.SheetNumber == sheetNum1Based)
                {
                    // Excel re-writes these to point to no sheet
                    nr.SheetNumber = (0);
                }
                else if (nr.SheetNumber > sheetNum1Based)
                {
                    // Bump down by one, so still points
                    //  at the same sheet
                    nr.SheetNumber = (nr.SheetNumber - 1);
                    // also update the link-table as otherwise references might point at invalid sheets
                }
            }
            if (linkTable != null)
            {
                // also tell the LinkTable about the removed sheet
                // +1 because we already removed it from the count of sheets!
                for (int i = sheetIndex + 1; i < NumSheets + 1; i++)
                {
                    linkTable.RemoveSheet(i);
                }
            }
        }

        /// <summary>
        /// make the tabid record look like the current situation.
        /// </summary>
        /// <returns>number of bytes written in the TabIdRecord</returns>
        private int FixTabIdRecord()
        {
            TabIdRecord tir = (TabIdRecord)records[records.Tabpos];
            int sz = tir.RecordSize;
            short[] tia = new short[boundsheets.Count];

            for (short k = 0; k < tia.Length; k++)
            {
                tia[k] = k;
            }
            tir.SetTabIdArray(tia);
            return tir.RecordSize - sz;
        }

        /**
         * returns the number of boundsheet objects contained in this workbook.
         *
         * @return number of BoundSheet records
         */

        public int NumSheets
        {
            get
            {
                //if (log.Check(POILogger.DEBUG))
                //    log.Log(DEBUG, "GetNumSheets=", boundsheets.Count);
                return boundsheets.Count;
            }
        }

        /**
         * Get the number of ExtendedFormat records contained in this workbook.
         *
         * @return int count of ExtendedFormat records
         */

        public int NumExFormats
        {
            get
            {
                //if (log.Check(POILogger.DEBUG))
                //    log.Log(DEBUG, "GetXF=", numxfs);
                return numxfs;
            }
        }
        /**
 * Retrieves the index of the given font
 */
        public int GetFontIndex(FontRecord font)
        {
            for (int i = 0; i <= numfonts; i++)
            {
                FontRecord thisFont =
                    (FontRecord)records[(records.Fontpos - (numfonts - 1)) + i];
                if (thisFont == font)
                {
                    // There is no 4!
                    if (i > 3)
                    {
                        return (i + 1);
                    }
                    return i;
                }
            }
            throw new ArgumentException("Could not find that font!");
        }
        /**
 * Returns the StyleRecord for the given
 *  xfIndex, or null if that ExtendedFormat doesn't
 *  have a Style set.
 */
        public StyleRecord GetStyleRecord(int xfIndex)
        {
            // Style records always follow after 
            //  the ExtendedFormat records
            bool done = false;
            for (int i = records.Xfpos; i < records.Count &&
                    !done; i++)
            {
                Record r = records[i];
                if (r is ExtendedFormatRecord)
                {
                }
                else if (r is StyleRecord)
                {
                    StyleRecord sr = (StyleRecord)r;
                    if (sr.XFIndex == xfIndex)
                    {
                        return sr;
                    }
                }
                else
                {
                    done = true;
                }
            }
            return null;
        }
        /**
         * Gets the ExtendedFormatRecord at the given 0-based index
         *
         * @param index of the Extended format record (0-based)
         * @return ExtendedFormatRecord at the given index
         */

        public ExtendedFormatRecord GetExFormatAt(int index)
        {
            int xfptr = records.Xfpos - (numxfs - 1);

            xfptr += index;
            ExtendedFormatRecord retval =
            (ExtendedFormatRecord)records[xfptr];

            return retval;
        }

        /**
         * Creates a new Cell-type Extneded Format Record and Adds it to the end of
         *  ExtendedFormatRecords collection
         *
         * @return ExtendedFormatRecord that was Created
         */

        public ExtendedFormatRecord CreateCellXF()
        {
            ExtendedFormatRecord xf = CreateExtendedFormat();
            records.Add(records.Xfpos + 1, xf);
            records.Xfpos=records.Xfpos + 1;
            numxfs++;
            return xf;
        }

        /**
         * Adds a string to the SST table and returns its index (if its a duplicate
         * just returns its index and update the counts) ASSUMES compressed Unicode
         * (meaning 8bit)
         *
         * @param string the string to be Added to the SSTRecord
         *
         * @return index of the string within the SSTRecord
         */

        public int AddSSTString(UnicodeString str)
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "Insert to sst string='", str);
            if (sst == null)
            {
                InsertSST();
            }
            return sst.AddString(str);
        }

        /**
         * given an index into the SST table, this function returns the corresponding String value
         * @return String containing the SST String
         */

        public UnicodeString GetSSTString(int str)
        {
            if (sst == null)
            {
                InsertSST();
            }
            UnicodeString retval = sst.GetString(str);

            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "Returning SST for index=", str,
                    //" String= ", retval);
            return retval;
        }

        /**
         * use this function to Add a Shared String Table to an existing sheet (say
         * generated by a different java api) without an sst....
         * @see #CreateSST()
         * @see org.apache.poi.hssf.record.SSTRecord
         */

        public void InsertSST()
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "creating new SST via InsertSST!");
            sst = new SSTRecord();
            records.Add(records.Count- 1, CreateExtendedSST());
            records.Add(records.Count - 2, sst);
        }

        /*
         * Serializes all records int the worksheet section into a big byte array. Use
         * this to Write the Workbook out.
         *
         * @return byte array containing the HSSF-only portions of the POIFS file.
         */
        // GJS: Not used so why keep it.
        //    public byte [] Serialize() {
        //        log.Log(DEBUG, "Serializing Workbook!");
        //        byte[] retval    = null;
        //
        ////         ArrayList bytes     = new ArrayList(records.Count);
        //        int    arraysize = Size;
        //        int    pos       = 0;
        //
        //        retval = new byte[ arraysize ];
        //        for (int k = 0; k < records.Count; k++) {
        //
        //            Record record = records[k];
        ////             Let's skip RECALCID records, as they are only use for optimization
        //	    if(record.Sid != RecalcIdRecord.Sid || ((RecalcIdRecord)record).IsNeeded) {
        //                pos += record.Serialize(pos, retval);   // rec.Length;
        //	    }
        //        }
        //        log.Log(DEBUG, "Exiting Serialize workbook");
        //        return retval;
        //    }

        /**
         * Serializes all records int the worksheet section into a big byte array. Use
         * this to Write the Workbook out.
         * @param offset of the data to be written
         * @param data array of bytes to Write this to
         */

        public int Serialize(int offset, byte[] data)
        {
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "Serializing Workbook with offsets");

            int pos = 0;

            SSTRecord sst = null;
            int sstPos = 0;
            bool wroteBoundSheets = false;
            for (int k = 0; k < records.Count; k++)
            {

                Record record = records[k];
                // Let's skip RECALCID records, as they are only use for optimization
                if (record.Sid != RecalcIdRecord.sid || ((RecalcIdRecord)record).IsNeeded)
                {
                    int len = 0;
                    if (record is SSTRecord)
                    {
                        sst = (SSTRecord)record;
                        sstPos = pos;
                    }
                    if (record.Sid == ExtSSTRecord.sid && sst != null)
                    {
                        record = sst.CreateExtSSTRecord(sstPos + offset);
                    }
                    if (record is BoundSheetRecord)
                    {
                        if (!wroteBoundSheets)
                        {
                            for (int i = 0; i < boundsheets.Count; i++)
                            {
                                len += ((BoundSheetRecord)boundsheets[i])
                                                 .Serialize(pos + offset + len, data);
                            }
                            wroteBoundSheets = true;
                        }
                    }
                    else
                    {
                        len = record.Serialize(pos + offset, data);
                    }
                    /////  DEBUG BEGIN /////
                    //                if (len != record.RecordSize)
                    //                    throw new InvalidOperationException("Record size does not match Serialized bytes.  Serialized size = " + len + " but RecordSize returns " + record.RecordSize);
                    /////  DEBUG END /////
                    pos += len;   // rec.Length;
                }
            }
            //if (log.Check(POILogger.DEBUG))
            //    log.Log(DEBUG, "Exiting Serialize workbook");
            return pos;
        }
        /**
         * Perform any work necessary before the workbook is about to be serialized.
         *
         * Include in it ant code that modifies the workbook record stream and affects its size.
         */
        public void PreSerialize()
        {
            // Ensure we have enough tab IDs
            // Can be a few short if new sheets were added
            if (records.Tabpos > 0)
            {
                TabIdRecord tir = (TabIdRecord)records[(records.Tabpos)];
                if (tir._tabids.Length < boundsheets.Count)
                {
                    FixTabIdRecord();
                }
            }
        }
        public int Size
        {
            get
            {
                int retval = 0;

                SSTRecord sst = null;
                for (int k = 0; k < records.Count; k++)
                {
                    Record record = records[k];
                    // Let's skip RECALCID records, as they are only use for optimization
                    if (record.Sid != RecalcIdRecord.sid || ((RecalcIdRecord)record).IsNeeded)
                    {
                        if (record is SSTRecord)
                            sst = (SSTRecord)record;
                        if (record.Sid == ExtSSTRecord.sid && sst != null)
                            retval += sst.CalcExtSSTRecordSize();
                        else
                            retval += record.RecordSize;
                    }
                }
                return retval;
            }
        }

        /**
         * Creates the BOF record
         * @see org.apache.poi.hssf.record.BOFRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a BOFRecord
         */

        private static Record CreateBOF()
        {
            BOFRecord retval = new BOFRecord();

            retval.Version=(short)0x600;
            retval.Type = BOFRecordType.Workbook;
            retval.Build=(short)0x10d3;

            //        retval.Build=(short)0x0dbb;
            retval.BuildYear=(short)1996;
            retval.HistoryBitMask=0x41;   // was c1 before verify
            retval.RequiredVersion=0x6;
            return retval;
        }

        /**
         * Creates the InterfaceHdr record
         * @see org.apache.poi.hssf.record.InterfaceHdrRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a InterfaceHdrRecord
         */
        [Obsolete]
        protected Record CreateInterfaceHdr()
        {
            //InterfaceHdrRecord retval = new InterfaceHdrRecord(CODEPAGE);

            //retval.Codepage=CODEPAGE;
            //return retval;
            return null;
        }

        /**
         * Creates an MMS record
         * @see org.apache.poi.hssf.record.MMSRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a MMSRecord
         */

        private static Record CreateMMS()
        {
            MMSRecord retval = new MMSRecord();

            retval.AddMenuCount=((byte)0);
            retval.DelMenuCount=((byte)0);
            return retval;
        }

        /**
         * Creates the InterfaceEnd record
         * @see org.apache.poi.hssf.record.InterfaceEndRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a InterfaceEndRecord
         */
        [Obsolete]
        protected Record CreateInterfaceEnd()
        {
            //return new InterfaceEndRecord();
            return null;
        }

        /**
         * Creates the WriteAccess record containing the logged in user's name
         * @see org.apache.poi.hssf.record.WriteAccessRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a WriteAccessRecord
         */

        private static Record CreateWriteAccess()
        {
            WriteAccessRecord retval = new WriteAccessRecord();
            String defaultUserName = "NPOI";
            try
            {
                String username = (Environment.UserName);
                // Google App engine returns null for user.name, see Bug 53974
                if (string.IsNullOrEmpty(username)) username = defaultUserName;

                retval.Username = (username);
            }
            catch (SecurityException)
            {
                // AccessControlException can occur in a restricted context
                // (client applet/jws application or restricted security server)
                retval.Username = (defaultUserName);
            }
            return retval;
        }

        /**
         * Creates the Codepage record containing the constant stored in CODEPAGE
         * @see org.apache.poi.hssf.record.CodepageRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a CodepageRecord
         */

        private static Record CreateCodepage()
        {
            CodepageRecord retval = new CodepageRecord();

            retval.Codepage=(CODEPAGE);
            return retval;
        }

        /**
         * Creates the DSF record containing a 0 since HSSF can't even Create Dual Stream Files
         * @see org.apache.poi.hssf.record.DSFRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a DSFRecord
         */

        private static Record CreateDSF()
        {
            return new DSFRecord(false); // we don't even support double stream files
        }

        /**
         * Creates the TabId record containing an array of 0,1,2.  This release of HSSF
         * always has the default three sheets, no less, no more.
         * @see org.apache.poi.hssf.record.TabIdRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a TabIdRecord
         */

        private static Record CreateTabId()
        {
            TabIdRecord retval = new TabIdRecord();
            short[] tabidarray = {
            0
        };

            retval.SetTabIdArray(tabidarray);
            return retval;
        }

        /**
         * Creates the FnGroupCount record containing the Magic number constant of 14.
         * @see org.apache.poi.hssf.record.FnGroupCountRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a FnGroupCountRecord
         */

        private static Record CreateFnGroupCount()
        {
            FnGroupCountRecord retval = new FnGroupCountRecord();

            retval.Count=(short)14;
            return retval;
        }

        /**
         * Creates the WindowProtect record with protect Set to false.
         * @see org.apache.poi.hssf.record.WindowProtectRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a WindowProtectRecord
         */

        private static Record CreateWindowProtect()
        {
            // by default even when we support it we won't
            // want it to be protected
            return new WindowProtectRecord(false);
        }

        /**
         * Creates the Protect record with protect Set to false.
         * @see org.apache.poi.hssf.record.ProtectRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a ProtectRecord
         */

        private static ProtectRecord CreateProtect()
        {
            return new ProtectRecord(false);
        }

        /**
         * Creates the Password record with password Set to 0.
         * @see org.apache.poi.hssf.record.PasswordRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a PasswordRecord
         */

        private static Record CreatePassword()
        {
            return new PasswordRecord(0x0000); // no password by default!
        }

        /**
         * Creates the ProtectionRev4 record with protect Set to false.
         * @see org.apache.poi.hssf.record.ProtectionRev4Record
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a ProtectionRev4Record
         */

        private static ProtectionRev4Record CreateProtectionRev4()
        {
            return new ProtectionRev4Record(false);
        }

        /**
         * Creates the PasswordRev4 record with password Set to 0.
         * @see org.apache.poi.hssf.record.PasswordRev4Record
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a PasswordRev4Record
         */

        private static Record CreatePasswordRev4()
        {
            return new PasswordRev4Record(0x0000);
        }

        /**
         * Creates the WindowOne record with the following magic values: 
         * horizontal hold - 0x168 
         * vertical hold   - 0x10e 
         * width           - 0x3a5c 
         * height          - 0x23be 
         * options         - 0x38 
         * selected tab    - 0 
         * Displayed tab   - 0 
         * num selected tab- 0 
         * tab width ratio - 0x258 
         * @see org.apache.poi.hssf.record.WindowOneRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a WindowOneRecord
         */

        private static Record CreateWindowOne()
        {
            WindowOneRecord retval = new WindowOneRecord();

            retval.HorizontalHold=(short)0x168;
            retval.VerticalHold=(short)0x10e;
            retval.Width=(short)0x3a5c;
            retval.Height=(short)0x23be;
            retval.Options=(short)0x38;
            retval.ActiveSheetIndex=(short)0x0;
            retval.FirstVisibleTab = (short)0x0;
            retval.NumSelectedTabs=(short)1;
            retval.TabWidthRatio=(short)0x258;
            return retval;
        }

        /**
         * Creates the Backup record with backup Set to 0. (loose the data, who cares)
         * @see org.apache.poi.hssf.record.BackupRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a BackupRecord
         */

        private static Record CreateBackup()
        {
            BackupRecord retval = new BackupRecord();

            retval.Backup=(short)0;   // by default DONT save backups of files...just loose data
            return retval;
        }

        /**
         * Creates the HideObj record with hide object Set to 0. (don't hide)
         * @see org.apache.poi.hssf.record.HideObjRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a HideObjRecord
         */

        private static Record CreateHideObj()
        {
            HideObjRecord retval = new HideObjRecord();

            retval.SetHideObj((short)0);   // by default Set hide object off
            return retval;
        }

        /**
         * Creates the DateWindow1904 record with windowing Set to 0. (don't window)
         * @see org.apache.poi.hssf.record.DateWindow1904Record
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a DateWindow1904Record
         */

        private static Record CreateDateWindow1904()
        {
            DateWindow1904Record retval = new DateWindow1904Record();

            retval.Windowing=((short)0);   // don't EVER use 1904 date windowing...tick tock..
            return retval;
        }

        /**
         * Creates the Precision record with precision Set to true. (full precision)
         * @see org.apache.poi.hssf.record.PrecisionRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a PrecisionRecord
         */

        private static Record CreatePrecision()
        {
            PrecisionRecord retval = new PrecisionRecord();

            retval.FullPrecision=(true);   // always use real numbers in calculations!
            return retval;
        }

        /**
         * Creates the RefreshAll record with refreshAll Set to true. (refresh all calcs)
         * @see org.apache.poi.hssf.record.RefreshAllRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a RefreshAllRecord
         */

        private static Record CreateRefreshAll()
        {
            return new RefreshAllRecord(false);
        }

        /**
         * Creates the BookBool record with saveLinkValues Set to 0. (don't save link values)
         * @see org.apache.poi.hssf.record.BookBoolRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a BookBoolRecord
         */

        private static Record CreateBookBool()
        {
            BookBoolRecord retval = new BookBoolRecord();

            retval.SaveLinkValues=(short)0;
            return retval;
        }

        /**
         * Creates a Font record with the following magic values: 
         * fontheight           = 0xc8
         * attributes           = 0x0
         * color palette index  = 0x7fff
         * bold weight          = 0x190
         * Font Name Length     = 5 
         * Font Name            = Arial 
         *
         * @see org.apache.poi.hssf.record.FontRecord
         * @see org.apache.poi.hssf.record.Record
         * @return record containing a FontRecord
         */

        private static Record CreateFont()
        {
            FontRecord retval = new FontRecord();

            retval.FontHeight=(short)0xc8;
            retval.Attributes=(short)0x0;
            retval.ColorPaletteIndex=(short)0x7fff;
            retval.BoldWeight=(short)0x190;
            retval.FontName="Arial";
            return retval;
        }

        // /**
        // * Creates a FormatRecord object
        // * @param id    the number of the format record to Create (meaning its position in
        // *        a file as M$ Excel would Create it.)
        // * @return record containing a FormatRecord
        // * @see org.apache.poi.hssf.record.FormatRecord
        // * @see org.apache.poi.hssf.record.Record
        // */
        
        //protected Record CreateFormat(int id)
        //{   // we'll need multiple editions for
        //    FormatRecord retval = new FormatRecord();   // the differnt formats

        //    switch (id)
        //    {

        //        case 0:
        //            retval.SetIndexCode((short)5);
        //            retval.SetFormatStringLength((byte)0x17);
        //            retval.SetFormatString("\"$\"#,##0_);\\(\"$\"#,##0\\)");
        //            break;

        //        case 1:
        //            retval.SetIndexCode((short)6);
        //            retval.SetFormatStringLength((byte)0x1c);
        //            retval.SetFormatString("\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)");
        //            break;

        //        case 2:
        //            retval.SetIndexCode((short)7);
        //            retval.SetFormatStringLength((byte)0x1d);
        //            retval.SetFormatString("\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)");
        //            break;

        //        case 3:
        //            retval.SetIndexCode((short)8);
        //            retval.SetFormatStringLength((byte)0x22);
        //            retval.SetFormatString(
        //            "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)");
        //            break;

        //        case 4:
        //            retval.SetIndexCode((short)0x2a);
        //            retval.SetFormatStringLength((byte)0x32);
        //            retval.SetFormatString(
        //            "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)");
        //            break;

        //        case 5:
        //            retval.SetIndexCode((short)0x29);
        //            retval.SetFormatStringLength((byte)0x29);
        //            retval.SetFormatString(
        //            "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)");
        //            break;

        //        case 6:
        //            retval.SetIndexCode((short)0x2c);
        //            retval.SetFormatStringLength((byte)0x3a);
        //            retval.SetFormatString(
        //            "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)");
        //            break;

        //        case 7:
        //            retval.SetIndexCode((short)0x2b);
        //            retval.SetFormatStringLength((byte)0x31);
        //            retval.SetFormatString(
        //            "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)");
        //            break;
        //    }
        //    return retval;
        //}

        /**
         * Creates an ExtendedFormatRecord object
         * @param id    the number of the extended format record to Create (meaning its position in
         *        a file as MS Excel would Create it.)
         *
         * @return record containing an ExtendedFormatRecord
         * @see org.apache.poi.hssf.record.ExtendedFormatRecord
         * @see org.apache.poi.hssf.record.Record
         */

        private static Record CreateExtendedFormat(int id)
        {   // we'll need multiple editions
            ExtendedFormatRecord retval = new ExtendedFormatRecord();

            switch (id)
            {

                case 0:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=(short)0;
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 1:
                    retval.FontIndex=(short)1;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 2:
                    retval.FontIndex=(short)1;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 3:
                    retval.FontIndex=(short)2;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 4:
                    retval.FontIndex=(short)2;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 5:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 6:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 7:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 8:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 9:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 10:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 11:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 12:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 13:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 14:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff400);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                // cell records
                case 15:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0;
                    retval.CellOptions=(short)0x1;
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=(short)0x0;
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                // style
                case 16:
                    retval.FontIndex=(short)1;
                    retval.FormatIndex=(short)0x2b;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff800);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 17:
                    retval.FontIndex=(short)1;
                    retval.FormatIndex=(short)0x29;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff800);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 18:
                    retval.FontIndex=(short)1;
                    retval.FormatIndex=(short)0x2c;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff800);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 19:
                    retval.FontIndex=(short)1;
                    retval.FormatIndex=(short)0x2a;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff800);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 20:
                    retval.FontIndex=(short)1;
                    retval.FormatIndex=(short)0x9;
                    retval.CellOptions=unchecked((short)0xfffffff5);
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=unchecked((short)0xfffff800);
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                // Unused from this point down
                case 21:
                    retval.FontIndex=(short)5;
                    retval.FormatIndex=(short)0x0;
                    retval.CellOptions=(short)0x1;
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=(short)0x800;
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 22:
                    retval.FontIndex=(short)6;
                    retval.FormatIndex=(short)0x0;
                    retval.CellOptions=(short)0x1;
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=(short)0x5c00;
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 23:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0x31;
                    retval.CellOptions=(short)0x1;
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=(short)0x5c00;
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 24:
                    retval.FontIndex=(short)0;
                    retval.FormatIndex=(short)0x8;
                    retval.CellOptions=(short)0x1;
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=(short)0x5c00;
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;

                case 25:
                    retval.FontIndex=(short)6;
                    retval.FormatIndex=(short)0x8;
                    retval.CellOptions=(short)0x1;
                    retval.AlignmentOptions=(short)0x20;
                    retval.IndentionOptions=(short)0x5c00;
                    retval.BorderOptions=(short)0;
                    retval.PaletteOptions=(short)0;
                    retval.AdtlPaletteOptions=(short)0;
                    retval.FillPaletteOptions=(short)0x20c0;
                    break;
            }
            return retval;
        }

        /**
         * Creates an default cell type ExtendedFormatRecord object.
         * @return ExtendedFormatRecord with intial defaults (cell-type)
         */

        private static ExtendedFormatRecord CreateExtendedFormat()
        {
            ExtendedFormatRecord retval = new ExtendedFormatRecord();

            retval.FontIndex=(short)0;
            retval.FormatIndex=(short)0x0;
            retval.CellOptions=(short)0x1;
            retval.AlignmentOptions=(short)0x20;
            retval.IndentionOptions=(short)0;
            retval.BorderOptions=(short)0;
            retval.PaletteOptions=(short)0;
            retval.AdtlPaletteOptions=(short)0;
            retval.FillPaletteOptions=(short)0x20c0;
            retval.TopBorderPaletteIdx=HSSFColor.Black.Index;
            retval.BottomBorderPaletteIdx=HSSFColor.Black.Index;
            retval.LeftBorderPaletteIdx=HSSFColor.Black.Index;
            retval.RightBorderPaletteIdx=HSSFColor.Black.Index;
            return retval;
        }

        /**
     * Creates a new StyleRecord, for the given Extended
     *  Format index, and adds it onto the end of the
     *  records collection
     */
        public StyleRecord CreateStyleRecord(int xfIndex) {
            // Style records always follow after 
            //  the ExtendedFormat records
            StyleRecord newSR = new StyleRecord();
            newSR.XFIndex = (short)xfIndex;
            
            // Find the spot
            int addAt = -1;
            for(int i=records.Xfpos; i<records.Count &&
                    addAt == -1; i++) {
                Record r = records[i];
                if(r is ExtendedFormatRecord ||
                        r is StyleRecord) {
                    // Keep going
                } else {
                    addAt = i;
                }
            }
            if(addAt == -1) {
                throw new InvalidOperationException("No XF Records found!");
            }
            records.Add(addAt, newSR);
            
            return newSR;
      }
        /**
         * Creates a StyleRecord object
         * @param id        the number of the style record to Create (meaning its position in
         *                  a file as MS Excel would Create it.
         * @return record containing a StyleRecord
         * @see org.apache.poi.hssf.record.StyleRecord
         * @see org.apache.poi.hssf.record.Record
         */

        private static Record CreateStyle(int id)
        {   // we'll need multiple editions
            StyleRecord retval = new StyleRecord();

            switch (id)
            {

                case 0:
                    retval.XFIndex = (unchecked((short)0xffff8010));
                    retval.SetBuiltinStyle(3);
                    retval.OutlineStyleLevel= (unchecked((byte)0xffffffff));
                    break;

                case 1:
                    retval.XFIndex = (unchecked((short)0xffff8011));
                    retval.SetBuiltinStyle(6);
                    retval.OutlineStyleLevel= (unchecked((byte)0xffffffff));
                    break;

                case 2:
                    retval.XFIndex = (unchecked((short)0xffff8012));
                    retval.SetBuiltinStyle(4);
                    retval.OutlineStyleLevel= (unchecked((byte)0xffffffff));
                    break;

                case 3:
                    retval.XFIndex = (unchecked((short)0xffff8013));
                    retval.SetBuiltinStyle(7);
                    retval.OutlineStyleLevel= (unchecked((byte)0xffffffff));
                    break;

                case 4:
                    retval.XFIndex = (unchecked((short)0xffff8000));
                    retval.SetBuiltinStyle(0);
                    retval.OutlineStyleLevel= (unchecked((byte)0xffffffff));
                    break;

                case 5:
                    retval.XFIndex = (unchecked((short)0xffff8014));
                    retval.SetBuiltinStyle(5);
                    retval.OutlineStyleLevel= (unchecked((byte)0xffffffff));
                    break;
            }
            return retval;
        }

        /**
         * Creates a palette record initialized to the default palette
         * @return a PaletteRecord instance populated with the default colors
         * @see org.apache.poi.hssf.record.PaletteRecord
         */
        private static PaletteRecord CreatePalette()
        {
            return new PaletteRecord();
        }

        /**
         * Creates the UseSelFS object with the use natural language flag Set to 0 (false)
         * @return record containing a UseSelFSRecord
         * @see org.apache.poi.hssf.record.UseSelFSRecord
         * @see org.apache.poi.hssf.record.Record
         */

        private static UseSelFSRecord CreateUseSelFS()
        {
            return new UseSelFSRecord(false);
        }

        /**
         * Create a "bound sheet" or "bundlesheet" (depending who you ask) record
         * Always Sets the sheet's bof to 0.  You'll need to Set that yourself.
         * @param id either sheet 0,1 or 2.
         * @return record containing a BoundSheetRecord
         * @see org.apache.poi.hssf.record.BoundSheetRecord
         * @see org.apache.poi.hssf.record.Record
         */

        private static Record CreateBoundSheet(int id)
        {   
            return new BoundSheetRecord("Sheet" + (id + 1));
        }

        /**
         * Creates the Country record with the default country Set to 1
         * and current country Set to 7 in case of russian locale ("ru_RU") and 1 otherwise
         * @return record containing a CountryRecord
         * @see org.apache.poi.hssf.record.CountryRecord
         * @see org.apache.poi.hssf.record.Record
         */

        private static Record CreateCountry()
        {   // what a novel idea, Create your own!
            CountryRecord retval = new CountryRecord();

            retval.DefaultCountry=((short)1);

            // from Russia with love ;)
            if (System.Threading.Thread.CurrentThread.CurrentCulture.Name.Equals("ru_RU"))
            {
                retval.CurrentCountry=((short)7);
            }
            else
            {
                retval.CurrentCountry=((short)1);
            }

            return retval;
        }

        /**
         * Creates the ExtendedSST record with numstrings per bucket Set to 0x8.  HSSF
         * doesn't yet know what to do with this thing, but we Create it with nothing in
         * it hardly just to make Excel happy and our sheets look like Excel's
         *
         * @return record containing an ExtSSTRecord
         * @see org.apache.poi.hssf.record.ExtSSTRecord
         * @see org.apache.poi.hssf.record.Record
         */

        private static Record CreateExtendedSST()
        {
            ExtSSTRecord retval = new ExtSSTRecord();

            retval.NumStringsPerBucket=((short)0x8);
            return retval;
        }

        /**
         * lazy initialization
         * Note - creating the link table causes creation of 1 EXTERNALBOOK and 1 EXTERNALSHEET record
         */
        private LinkTable OrCreateLinkTable
        {
            get
            {
                return GetOrCreateLinkTable();
            }
        }

        private LinkTable GetOrCreateLinkTable()
        {
            if (linkTable == null)
            {
                linkTable = new LinkTable((short)NumSheets, records);
            }
            return linkTable;
        }

        public int LinkExternalWorkbook(String name, IWorkbook externalWorkbook)
        {
            return GetOrCreateLinkTable().LinkExternalWorkbook(name, externalWorkbook);
        }

        /** 
         * Finds the first sheet name by his extern sheet index
         * @param externSheetIndex extern sheet index
         * @return first sheet name.
         */
        public String FindSheetFirstNameFromExternSheet(int externSheetIndex)
        {
            int indexToSheet = linkTable.GetFirstInternalSheetIndexForExtIndex(externSheetIndex);
            return FindSheetNameFromIndex(indexToSheet);
        }
        public String FindSheetLastNameFromExternSheet(int externSheetIndex)
        {
            int indexToSheet = linkTable.GetLastInternalSheetIndexForExtIndex(externSheetIndex);
            return FindSheetNameFromIndex(indexToSheet);
        }
        private String FindSheetNameFromIndex(int internalSheetIndex)
        {
            if (internalSheetIndex < 0)
            {
                // TODO - what does '-1' mean here?
                //error Check, bail out gracefully!
                return "";
            }
            if (internalSheetIndex >= boundsheets.Count)
            {
                // Not sure if this can ever happen (See bug 45798)
                return ""; // Seems to be what excel would do in this case
            }
            return GetSheetName(internalSheetIndex);
        }

        public ExternalSheet GetExternalSheet(int externSheetIndex)
        {
            String[] extNames = linkTable.GetExternalBookAndSheetName(externSheetIndex);
            if (extNames == null)
            {
                return null;
            }
            if (extNames.Length == 2)
            {
                return new ExternalSheet(extNames[0], extNames[1]);
            }
            else
            {
                return new ExternalSheetRange(extNames[0], extNames[1], extNames[2]);
            }
        }


        /**
         * Finds the (first) sheet index for a particular external sheet number.
         * @param externSheetNumber     The external sheet number to convert
         * @return  The index to the sheet found.
         */
        public int GetFirstSheetIndexFromExternSheetIndex(int externSheetNumber)
        {
            return linkTable.GetFirstInternalSheetIndexForExtIndex(externSheetNumber);
        }

        /**
         * Finds the last sheet index for a particular external sheet number,
         *  which may be the same as the first (except for multi-sheet references)
         * @param externSheetNumber     The external sheet number to convert
         * @return  The index to the sheet found.
         */
        public int GetLastSheetIndexFromExternSheetIndex(int externSheetNumber)
        {
            return linkTable.GetLastInternalSheetIndexForExtIndex(externSheetNumber);
        }

        /** 
         * Returns the extern sheet number for specific sheet number.
         * If this sheet doesn't exist in extern sheet, add it
         * @param sheetNumber local sheet number
         * @return index to extern sheet
         */
        public int CheckExternSheet(int sheetNumber)
        {
            return OrCreateLinkTable.CheckExternSheet(sheetNumber);
        }
        /** 
         * Returns the extern sheet number for specific range of sheets.
         * If this sheet range doesn't exist in extern sheet, add it
         * @param firstSheetNumber first local sheet number
         * @param lastSheetNumber last local sheet number
         * @return index to extern sheet
         */
        public short checkExternSheet(int firstSheetNumber, int lastSheetNumber)
        {
            return (short)OrCreateLinkTable.CheckExternSheet(firstSheetNumber, lastSheetNumber);
        }

        public int GetExternalSheetIndex(String workbookName, String sheetName)
        {
            return OrCreateLinkTable.GetExternalSheetIndex(workbookName, sheetName, sheetName);
        }
        public int GetExternalSheetIndex(String workbookName, String firstSheetName, String lastSheetName)
        {
            return OrCreateLinkTable.GetExternalSheetIndex(workbookName, firstSheetName, lastSheetName);
        }
        /** Gets the total number of names
         * @return number of names
         */
        public int NumNames
        {
            get
            {
                if (linkTable == null)
                {
                    return 0;
                }
                return linkTable.NumNames;
            }
        }
        /**
         *
         * @param name the  name of an external function, typically a name of a UDF
         * @param sheetRefIndex the sheet ref index, or -1 if not known
         * @param udf  locator of user-defiend functions to resolve names of VBA and Add-In functions
         * @return the external name or null
         */
        public NameXPtg GetNameXPtg(String name, int sheetRefIndex, UDFFinder udf)
        {
            LinkTable lnk = OrCreateLinkTable;
            NameXPtg xptg = lnk.GetNameXPtg(name, sheetRefIndex);

            if (xptg == null && udf.FindFunction(name) != null)
            {
                // the name was not found in the list of external names
                // check if the Workbook's UDFFinder is aware about it and register the name if it is
                xptg = lnk.AddNameXPtg(name);
            }
            return xptg;
        }
        public NameXPtg GetNameXPtg(String name, UDFFinder udf)
        {
            return GetNameXPtg(name, -1, udf);
        }
        /** Gets the name record
         * @param index name index
         * @return name record
         */
        public NameRecord GetNameRecord(int index)
        {
            return linkTable.GetNameRecord(index);
        }

        /** Creates new name
         * @return new name record
         */
        public NameRecord CreateName()
        {
            return AddName(new NameRecord());
        }


        /** Creates new name
         * @return new name record
         */
        public NameRecord AddName(NameRecord name)
        {

            OrCreateLinkTable.AddName(name);

            return name;
        }

        /**Generates a NameRecord to represent a built-in region
         * @return a new NameRecord Unless the index is invalid
         */
        public NameRecord CreateBuiltInName(byte builtInName, int index)
        {
            if (index == -1 || index + 1 > short.MaxValue)
                throw new ArgumentException("Index is not valid [" + index + "]");

            NameRecord name = new NameRecord(builtInName, (short)(index));

            AddName(name);

            return name;
        }


        /** Removes the name
         * @param namenum name index
         */
        public void RemoveName(int namenum)
        {

            if (linkTable.NumNames > namenum)
            {
                int idx = FindFirstRecordLocBySid(NameRecord.sid);
                records.Remove(idx + namenum);
                linkTable.RemoveName(namenum);
            }

        }
        private Dictionary<String, NameCommentRecord> commentRecords;
        /**
         * If a {@link NameCommentRecord} is added or the name it references
         *  is renamed, then this will update the lookup cache for it.
         */
        public void UpdateNameCommentRecordCache(NameCommentRecord commentRecord)
        {
            if (commentRecords.ContainsValue(commentRecord))
            {
                foreach (KeyValuePair<string, NameCommentRecord> entry in commentRecords)
                {
                    if (entry.Value.Equals(commentRecord))
                    {
                        commentRecords.Remove(entry.Key);
                        break;
                    }
                }
            }
            commentRecords[commentRecord.NameText] = commentRecord;
        }
        /**
         * Returns a format index that matches the passed in format.  It does not tie into HSSFDataFormat.
         * @param format the format string
         * @param CreateIfNotFound Creates a new format if format not found
         * @return the format id of a format that matches or -1 if none found and CreateIfNotFound
         */
        public short GetFormat(String format, bool CreateIfNotFound)
        {
            IEnumerator iterator;
            for (iterator = formats.GetEnumerator(); iterator.MoveNext(); )
            {
                FormatRecord r = (FormatRecord)iterator.Current;
                if (r.FormatString.Equals(format))
                {
                    return (short)r.IndexCode;
                }
            }

            if (CreateIfNotFound)
            {
                return (short)CreateFormat(format);
            }

            return -1;
        }

        /**
         * Returns the list of FormatRecords in the workbook.
         * @return ArrayList of FormatRecords in the notebook
         */
        public List<FormatRecord> Formats
        {
            get
            {
                return formats;
            }
        }

        /**
         * Creates a FormatRecord, Inserts it, and returns the index code.
         * @param format the format string
         * @return the index code of the format record.
         * @see org.apache.poi.hssf.record.FormatRecord
         * @see org.apache.poi.hssf.record.Record
         */
        public int CreateFormat(String formatString)
        {
            //        ++xfpos;	//These are to Ensure that positions are updated properly
            //        ++palettepos;
            //        ++bspos;
            maxformatid = maxformatid >= (short)0xa4 ? (short)(maxformatid + 1) : (short)0xa4; //Starting value from M$ empiracle study.
            FormatRecord rec = new FormatRecord(maxformatid, formatString);

            int pos = 0;
            while (pos < records.Count && records[pos].Sid != FormatRecord.sid)
                pos++;
            pos += formats.Count;
            formats.Add(rec);
            records.Add(pos, rec);
            return maxformatid;
        }

        /**
     * Creates a FormatRecord object
     * @param id    the number of the format record to create (meaning its position in
     *        a file as M$ Excel would create it.)
     */
        private static FormatRecord CreateFormat(int id)
        {
            // we'll need multiple editions for
            // the different formats


            switch (id)
            {
                case 0: return new FormatRecord(5, BuiltinFormats.GetBuiltinFormat(5));
                case 1: return new FormatRecord(6, BuiltinFormats.GetBuiltinFormat(6));
                case 2: return new FormatRecord(7, BuiltinFormats.GetBuiltinFormat(7));
                case 3: return new FormatRecord(8, BuiltinFormats.GetBuiltinFormat(8));
                case 4: return new FormatRecord(0x2a, BuiltinFormats.GetBuiltinFormat(0x2a));
                case 5: return new FormatRecord(0x29, BuiltinFormats.GetBuiltinFormat(0x29));
                case 6: return new FormatRecord(0x2c, BuiltinFormats.GetBuiltinFormat(0x2c));
                case 7: return new FormatRecord(0x2b, BuiltinFormats.GetBuiltinFormat(0x2b));
            }
            throw new ArgumentException("Unexpected id " + id);
        }

        /**
         * Returns the first occurance of a record matching a particular sid.
         */
        public Record FindFirstRecordBySid(short sid)
        {
            for (IEnumerator iterator = records.GetEnumerator(); iterator.MoveNext(); )
            {
                Record record = (Record)iterator.Current;

                if (record.Sid == sid)
                {
                    return record;
                }
            }
            return null;
        }

        /**
         * Returns the index of a record matching a particular sid.
         * @param sid   The sid of the record to match
         * @return      The index of -1 if no match made.
         */
        public int FindFirstRecordLocBySid(short sid)
        {
            int index = 0;
            for (IEnumerator iterator = records.GetEnumerator(); iterator.MoveNext(); )
            {
                Record record = (Record)iterator.Current;

                if (record.Sid == sid)
                {
                    return index;
                }
                index++;
            }
            return -1;
        }

        /**
         * Returns the next occurance of a record matching a particular sid.
         */
        public Record FindNextRecordBySid(short sid, int pos)
        {
            int matches = 0;
            for (IEnumerator iterator = records.GetEnumerator(); iterator.MoveNext(); )
            {
                Record record = (Record)iterator.Current;

                if (record.Sid == sid)
                {
                    if (matches++ == pos)
                        return record;
                }
            }
            return null;
        }

        public IList Hyperlinks
        {
            get{return hyperlinks;}
        }

        public IList Records
        {
            get{return records.Records;}
        }

        //    public void InsertChartRecords( List chartRecords )
        //    {
        //        backuppos += chartRecords.Count;
        //        fontpos += chartRecords.Count;
        //        palettepos += chartRecords.Count;
        //        bspos += chartRecords.Count;
        //        xfpos += chartRecords.Count;
        //
        //        records.AddAll(protpos, chartRecords);
        //    }

        /**
        * Whether date windowing is based on 1/2/1904 or 1/1/1900.
        * Some versions of Excel (Mac) can save workbooks using 1904 date windowing.
        *
        * @return true if using 1904 date windowing
        */
        public bool IsUsing1904DateWindowing
        {
            get { return uses1904datewindowing; }
        }

        /**
         * Returns the custom palette in use for this workbook; if a custom palette record
         * does not exist, then it is Created.
         */
        public PaletteRecord CustomPalette
        {
            get
            {
                PaletteRecord palette;
                int palettePos = records.Palettepos;
                if (palettePos != -1)
                {
                    Record rec = records[palettePos];
                    if (rec is PaletteRecord)
                    {
                        palette = (PaletteRecord)rec;
                    }
                    else throw new Exception("InternalError: Expected PaletteRecord but got a '" + rec + "'");
                }
                else
                {
                    palette = CreatePalette();
                    //Add the palette record after the bof which is always the first record
                    records.Add(1, palette);
                    records.Palettepos = 1;
                }
                return palette;
            }
        }

        /**
         * Finds the primary drawing Group, if one already exists
         */
        public DrawingManager2 FindDrawingGroup()
        {
            if (drawingManager != null)
            {
                // We already have it!
                return drawingManager;
            }

            // Need to Find a DrawingGroupRecord that
            //  Contains a EscherDggRecord
            for (IEnumerator rit = records.GetEnumerator(); rit.MoveNext(); )
            {
                Record r = (Record)rit.Current;

                if (r is DrawingGroupRecord)
                {
                    DrawingGroupRecord dg = (DrawingGroupRecord)r;
                    dg.ProcessChildRecords();

                    EscherContainerRecord cr =
                        dg.GetEscherContainer();
                    if (cr == null)
                    {
                        continue;
                    }

                    EscherDggRecord dgg = null;
                    EscherContainerRecord bStore = null;
                    for (IEnumerator it = cr.ChildRecords.GetEnumerator(); it.MoveNext(); )
                    {
                        EscherRecord er = (EscherRecord)it.Current;
                        if (er is EscherDggRecord)
                        {
                            dgg = (EscherDggRecord)er;
                        }
                        else if (er.RecordId == EscherContainerRecord.BSTORE_CONTAINER)
                        {
                            bStore = (EscherContainerRecord)er;
                        }
                    }

                    if (dgg != null)
                    {
                        drawingManager = new DrawingManager2(dgg);
                        if (bStore != null)
                        {
                            foreach (EscherRecord bs in bStore.ChildRecords)
                            {
                                if (bs is EscherBSERecord)
                                    escherBSERecords.Add((EscherBSERecord)bs);
                            }
                        }
                        return drawingManager;
                    }
                }
            }

            // Look for the DrawingGroup record
            int dgLoc = FindFirstRecordLocBySid(DrawingGroupRecord.sid);

            // If there is one, does it have a EscherDggRecord?
            if (dgLoc != -1)
            {
                DrawingGroupRecord dg =
                    (DrawingGroupRecord)records[dgLoc];
                EscherDggRecord dgg = null;
                EscherContainerRecord bStore = null;

                for (IEnumerator it = dg.EscherRecords.GetEnumerator(); it.MoveNext(); )
                {
                    EscherRecord er = (EscherRecord)it.Current;
                    if (er is EscherDggRecord)
                    {
                        dgg = (EscherDggRecord)er;
                    }
                    else if (er.RecordId == EscherContainerRecord.BSTORE_CONTAINER)
                    {
                        bStore = (EscherContainerRecord)er;
                    }
                }

                if (dgg != null)
                {
                    drawingManager = new DrawingManager2(dgg);
                    if (bStore != null)
                    {
                        foreach (EscherRecord bs in bStore.ChildRecords)
                        {
                            if (bs is EscherBSERecord)
                                escherBSERecords.Add((EscherBSERecord)bs);
                        }
                    }
                }
            }
            return drawingManager;
        }

        /**
         * Creates a primary drawing Group record.  If it already 
         *  exists then it's modified.
         */
        public void CreateDrawingGroup()
        {
            if (drawingManager == null)
            {
                EscherContainerRecord dggContainer = new EscherContainerRecord();
                EscherDggRecord dgg = new EscherDggRecord();
                EscherOptRecord opt = new EscherOptRecord();
                EscherSplitMenuColorsRecord splitMenuColors = new EscherSplitMenuColorsRecord();

                dggContainer.RecordId=unchecked((short)0xF000);
                dggContainer.Options=(short)0x000F;
                dgg.RecordId=EscherDggRecord.RECORD_ID;
                dgg.Options=(short)0x0000;
                dgg.ShapeIdMax=1024;
                dgg.NumShapesSaved=0;
                dgg.DrawingsSaved=0;
                dgg.FileIdClusters=new EscherDggRecord.FileIdCluster[] { };
                drawingManager = new DrawingManager2(dgg);
                EscherContainerRecord bstoreContainer = null;
                if (escherBSERecords.Count > 0)
                {
                    bstoreContainer = new EscherContainerRecord();
                    bstoreContainer.RecordId=EscherContainerRecord.BSTORE_CONTAINER;
                    bstoreContainer.Options=(short)((escherBSERecords.Count << 4) | 0xF);
                    for (IEnumerator iterator = escherBSERecords.GetEnumerator(); iterator.MoveNext(); )
                    {
                        EscherRecord escherRecord = (EscherRecord)iterator.Current;
                        bstoreContainer.AddChildRecord(escherRecord);
                    }
                }
                opt.RecordId=unchecked((short)0xF00B);
                opt.Options=(short)0x0033;
                opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 524296));
                opt.AddEscherProperty(new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, 0x08000041));
                opt.AddEscherProperty(new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, 134217792));
                splitMenuColors.RecordId=unchecked((short)0xF11E);
                splitMenuColors.Options=(short)0x0040;
                splitMenuColors.Color1=0x0800000D;
                splitMenuColors.Color2=0x0800000C;
                splitMenuColors.Color3=0x08000017;
                splitMenuColors.Color4=0x100000F7;

                dggContainer.AddChildRecord(dgg);
                if (bstoreContainer != null)
                    dggContainer.AddChildRecord(bstoreContainer);
                dggContainer.AddChildRecord(opt);
                dggContainer.AddChildRecord(splitMenuColors);

                int dgLoc = FindFirstRecordLocBySid(DrawingGroupRecord.sid);
                if (dgLoc == -1)
                {
                    DrawingGroupRecord drawingGroup = new DrawingGroupRecord();
                    drawingGroup.AddEscherRecord(dggContainer);
                    int loc = FindFirstRecordLocBySid(CountryRecord.sid);

                    Records.Insert(loc + 1, drawingGroup);
                }
                else
                {
                    DrawingGroupRecord drawingGroup = new DrawingGroupRecord();
                    drawingGroup.AddEscherRecord(dggContainer);
                    Records[dgLoc]= drawingGroup;
                }

            }
        }

        public WindowOneRecord WindowOne
        {
            get
            {
                return windowOne;
            }
        }

        /**
         * Removes the given font record from the
         *  file's list. This will make all 
         *  subsequent font indicies drop by one,
         *  so you'll need to update those yourself!
         */
        public void RemoveFontRecord(FontRecord rec)
        {
            records.Remove(rec); // this updates FontPos for us
            numfonts--;
        }
        /**
 * Removes the given ExtendedFormatRecord record from the
 *  file's list. This will make all 
 *  subsequent font indicies drop by one,
 *  so you'll need to update those yourself!
 */
        public void RemoveExFormatRecord(ExtendedFormatRecord rec)
        {
            records.Remove(rec); // this updates XfPos for us
            numxfs--;
        }

        /// <summary>
        /// Removes ExtendedFormatRecord record with given index from the file's list. This will make all
        /// subsequent font indicies drop by one,so you'll need to update those yourself!
        /// </summary>
        /// <param name="index">index of the Extended format record (0-based)</param>
        public void RemoveExFormatRecord(int index)
        {
            int xfptr = records.Xfpos - (numxfs - 1) + index;
            records.Remove(xfptr); // this updates XfPos for us
            numxfs--;
        }
        public EscherBSERecord GetBSERecord(int pictureIndex)
        {
            return (EscherBSERecord)escherBSERecords[pictureIndex - 1];
        }

        public int AddBSERecord(EscherBSERecord e)
        {
            CreateDrawingGroup();

            // maybe we don't need that as an instance variable anymore
            escherBSERecords.Add(e);

            int dgLoc = FindFirstRecordLocBySid(DrawingGroupRecord.sid);
            DrawingGroupRecord drawingGroup = (DrawingGroupRecord)Records[dgLoc];

            EscherContainerRecord dggContainer = (EscherContainerRecord)drawingGroup.GetEscherRecord(0);
            EscherContainerRecord bstoreContainer;
            if (dggContainer.GetChild(1).RecordId == EscherContainerRecord.BSTORE_CONTAINER)
            {
                bstoreContainer = (EscherContainerRecord)dggContainer.GetChild(1);
            }
            else
            {
                bstoreContainer = new EscherContainerRecord();
                bstoreContainer.RecordId=EscherContainerRecord.BSTORE_CONTAINER;

                //dggContainer.ChildRecords.Insert(1, bstoreContainer);
                List<EscherRecord> childRecords = dggContainer.ChildRecords;
                childRecords.Insert(1, bstoreContainer);
                dggContainer.ChildRecords = (childRecords);
                
            }
            bstoreContainer.Options=(short)((escherBSERecords.Count << 4) | 0xF);

            bstoreContainer.AddChildRecord(e);

            return escherBSERecords.Count;
        }

        public DrawingManager2 DrawingManager
        {
            get
            {
                return drawingManager;
            }
        }

        public WriteProtectRecord WriteProtect
        {
            get
            {
                if (this.writeProtect == null)
                {
                    this.writeProtect = new WriteProtectRecord();
                    int i = 0;
                    for (i = 0;
                         i < records.Count && !(records[i] is BOFRecord);
                         i++)
                    {
                    }
                    records.Add(i + 1, this.writeProtect);
                }
                return this.writeProtect;
            }
        }

        public WriteAccessRecord WriteAccess
        {
            get
            {
                if (this.writeAccess == null)
                {
                    this.writeAccess = (WriteAccessRecord)CreateWriteAccess();
                    int i = 0;
                    for (i = 0;
                         i < records.Count && !(records[i] is InterfaceEndRecord);
                         i++)
                    {
                    }
                    records.Add(i + 1, this.writeAccess);
                }
                return this.writeAccess;
            }
        }

        public FileSharingRecord FileSharing
        {
            get
            {
                if (this.fileShare == null)
                {
                    this.fileShare = new FileSharingRecord();
                    int i = 0;
                    for (i = 0;
                         i < records.Count && !(records[i] is WriteAccessRecord);
                         i++)
                    {
                    }
                    records.Add(i + 1, this.fileShare);
                }
                return this.fileShare;
            }
        }

        /**
         * is the workbook protected with a password (not encrypted)?
         */
        public bool IsWriteProtected
        {
            get
            {
                if (this.fileShare == null)
                {
                    return false;
                }
                FileSharingRecord frec = FileSharing;
                return (frec.ReadOnly == 1);
            }
        }

        /**
         * protect a workbook with a password (not encypted, just Sets Writeprotect
         * flags and the password.
         * @param password to Set
         */
        public void WriteProtectWorkbook(String password, String username)
        {
            FileSharingRecord frec = FileSharing;
            WriteAccessRecord waccess = WriteAccess;
            frec.ReadOnly=((short)1);
            frec.Password=(FileSharingRecord.HashPassword(password));
            frec.Username=(username);
            waccess.Username=(username);
        }

        /**
         * Removes the Write protect flag
         */
        public void UnwriteProtectWorkbook()
        {
            records.Remove(fileShare);
            records.Remove(WriteProtect);
            fileShare = null;
            writeProtect = null;
        }

        /**
         * @param reFindex Index to REF entry in EXTERNSHEET record in the Link Table
         * @param definedNameIndex zero-based to DEFINEDNAME or EXTERNALNAME record
         * @return the string representation of the defined or external name
         */
        public String ResolveNameXText(int reFindex, int definedNameIndex)
        {
            return linkTable.ResolveNameXText(reFindex, definedNameIndex, this);
        }

        public NameRecord CloneFilter(int filterDbNameIndex, int newSheetIndex)
        {
            NameRecord origNameRecord = GetNameRecord(filterDbNameIndex);
            // copy original formula but adjust 3D refs to the new external sheet index
            int newExtSheetIx = CheckExternSheet(newSheetIndex);
            Ptg[] ptgs = origNameRecord.NameDefinition;
            for (int i = 0; i < ptgs.Length; i++)
            {
                Ptg ptg = ptgs[i];
                if (ptg is Area3DPtg)
                {
                    Area3DPtg a3p = (Area3DPtg)((OperandPtg)ptg).Copy();
                    a3p.ExternSheetIndex = (newExtSheetIx);
                    ptgs[i] = a3p;
                }
                else if (ptg is Ref3DPtg)
                {
                    Ref3DPtg r3p = (Ref3DPtg)((OperandPtg)ptg).Copy();
                    r3p.ExternSheetIndex = (newExtSheetIx);
                    ptgs[i] = r3p;
                }
            }
            NameRecord newNameRecord = CreateBuiltInName(NameRecord.BUILTIN_FILTER_DB, newSheetIndex + 1);
            newNameRecord.NameDefinition = ptgs;
            newNameRecord.IsHiddenName = true;
            return newNameRecord;
        }

        /**
 * Updates named ranges due to moving of cells
 */
        public void UpdateNamesAfterCellShift(FormulaShifter shifter)
        {
            for (int i = 0; i < NumNames; ++i)
            {
                NameRecord nr = GetNameRecord(i);
                Ptg[] ptgs = nr.NameDefinition;
                if (shifter.AdjustFormula(ptgs, nr.SheetNumber))
                {
                    nr.NameDefinition = ptgs;
                }
            }
        }
        /**
     * Get or create RecalcIdRecord
     *
     * @see org.apache.poi.hssf.usermodel.HSSFWorkbook#setForceFormulaRecalculation(boolean)
     */
        public RecalcIdRecord RecalcId
        {
            get
            {
                RecalcIdRecord record = (RecalcIdRecord)FindFirstRecordBySid(RecalcIdRecord.sid);
                if (record == null)
                {
                    record = new RecalcIdRecord();
                    // typically goes after the Country record
                    int pos = FindFirstRecordLocBySid(CountryRecord.sid);
                    records.Add(pos + 1, record);
                }
                return record;
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
            return linkTable.ChangeExternalReference(oldUrl, newUrl);
        }
    }
}
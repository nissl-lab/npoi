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
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;
    using NPOI.SS.UserModel;

    /**
     * Link Table (OOO pdf reference: 4.10.3 ) <p/>
     *
     * The main data of all types of references is stored in the Link Table inside the Workbook Globals
     * Substream (4.2.5). The Link Table itself is optional and occurs only, if  there are any
     * references in the document.
     *  <p/>
     *
     *  In BIFF8 the Link Table consists of
     *  <ul>
     *  <li>zero or more EXTERNALBOOK Blocks<p/>
     *  	each consisting of
     *  	<ul>
     *  	<li>exactly one EXTERNALBOOK (0x01AE) record</li>
     *  	<li>zero or more EXTERNALNAME (0x0023) records</li>
     *  	<li>zero or more CRN Blocks<p/>
     *			each consisting of
     *  		<ul>
     *  		<li>exactly one XCT (0x0059)record</li>
     *  		<li>zero or more CRN (0x005A) records (documentation says one or more)</li>
     *  		</ul>
     *  	</li>
     *  	</ul>
     *  </li>
     *  <li>zero or one EXTERNSHEET (0x0017) record</li>
     *  <li>zero or more DEFINEDNAME (0x0018) records</li>
     *  </ul>
     *
     *
     * @author Josh Micich
     */
    public class LinkTable
    {
        // TODO make this class into a record aggregate

        private static ExternSheetRecord ReadExtSheetRecord(RecordStream rs)
        {
            List<ExternSheetRecord> temp = new List<ExternSheetRecord>(2);
            while (rs.PeekNextClass() == typeof(ExternSheetRecord))
            {
                temp.Add((ExternSheetRecord)rs.GetNext());
            }

            int nItems = temp.Count;
            if (nItems < 1)
            {
                throw new Exception("Expected an EXTERNSHEET record but got ("
                        + rs.PeekNextClass().Name + ")");
            }
            if (nItems == 1)
            {
                // this is the normal case. There should be just one ExternSheetRecord
                return temp[0];
            }
            // Some apps generate multiple ExternSheetRecords (see bug 45698).
            // It seems like the best thing to do might be to combine these into one
            ExternSheetRecord[] esrs = new ExternSheetRecord[nItems];
            esrs = temp.ToArray();
            return ExternSheetRecord.Combine(esrs);
        }

        private class CRNBlock
        {

            private CRNCountRecord _countRecord;
            private CRNRecord[] _crns;

            public CRNBlock(RecordStream rs)
            {
                _countRecord = (CRNCountRecord)rs.GetNext();
                int nCRNs = _countRecord.NumberOfCRNs;
                CRNRecord[] crns = new CRNRecord[nCRNs];
                for (int i = 0; i < crns.Length; i++)
                {
                    crns[i] = (CRNRecord)rs.GetNext();
                }
                _crns = crns;
            }
            public CRNRecord[] GetCrns()
            {
                return (CRNRecord[])_crns.Clone();
            }
        }

        private class ExternalBookBlock
        {
            private SupBookRecord _externalBookRecord;
            internal ExternalNameRecord[] _externalNameRecords;
            private CRNBlock[] _crnBlocks;
            /**
             * Create a new block for registering add-in functions
             *
             * @see org.apache.poi.hssf.model.LinkTable#addNameXPtg(String)
             */
            public ExternalBookBlock()
            {
                _externalBookRecord = SupBookRecord.CreateAddInFunctions();
                _externalNameRecords = new ExternalNameRecord[0];
                _crnBlocks = new CRNBlock[0];
            }
            public ExternalBookBlock(RecordStream rs)
            {
                _externalBookRecord = (SupBookRecord)rs.GetNext();
                ArrayList temp = new ArrayList();
                while (rs.PeekNextClass() == typeof(ExternalNameRecord))
                {
                    temp.Add(rs.GetNext());
                }
                _externalNameRecords = (ExternalNameRecord[])temp.ToArray(typeof(ExternalNameRecord));

                temp.Clear();

                while (rs.PeekNextClass() == typeof(CRNCountRecord))
                {
                    temp.Add(new CRNBlock(rs));
                }
                _crnBlocks = (CRNBlock[])temp.ToArray(typeof(CRNBlock));
            }

            /**
            * Create a new block for external references.
            */
            public ExternalBookBlock(String url, String[] sheetNames)
            {
                _externalBookRecord = SupBookRecord.CreateExternalReferences(url, sheetNames);
                _crnBlocks = new CRNBlock[0];
            }

            /**
             * Create a new block for internal references. It is called when constructing a new LinkTable.
             *
             * @see org.apache.poi.hssf.model.LinkTable#LinkTable(int, WorkbookRecordList)
             */
            public ExternalBookBlock(int numberOfSheets)
            {
                _externalBookRecord = SupBookRecord.CreateInternalReferences((short)numberOfSheets);
                _externalNameRecords = new ExternalNameRecord[0];
                _crnBlocks = new CRNBlock[0];
            }

            public int NumberOfNames
            {
                get
                {
                    return _externalNameRecords.Length;
                }
            }
            public int AddExternalName(ExternalNameRecord rec)
            {
                ExternalNameRecord[] tmp = new ExternalNameRecord[_externalNameRecords.Length + 1];
                Array.Copy(_externalNameRecords, 0, tmp, 0, _externalNameRecords.Length);
                tmp[tmp.Length - 1] = rec;
                _externalNameRecords = tmp;
                return _externalNameRecords.Length - 1;
            }


            public SupBookRecord GetExternalBookRecord()
            {
                return _externalBookRecord;
            }

            public String GetNameText(int definedNameIndex)
            {
                return _externalNameRecords[definedNameIndex].Text;
            }
            /**
             * Performs case-insensitive search
             * @return -1 if not found
             */
            public int GetIndexOfName(String name)
            {
                for (int i = 0; i < _externalNameRecords.Length; i++)
                {
                    if (_externalNameRecords[i].Text.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        return i;
                    }
                }
                return -1;
            }

            public int GetNameIx(int definedNameIndex)
            {
                return _externalNameRecords[definedNameIndex].Ix;
            }

        }

        private ExternalBookBlock[] _externalBookBlocks;
        private ExternSheetRecord _externSheetRecord;
        private List<NameRecord> _definedNames;
        private int _recordCount;
        private WorkbookRecordList _workbookRecordList; // TODO - would be nice to Remove this

        public LinkTable(List<Record> inputList, int startIndex, WorkbookRecordList workbookRecordList, Dictionary<String, NameCommentRecord> commentRecords)
        {

            _workbookRecordList = workbookRecordList;
            RecordStream rs = new RecordStream(inputList, startIndex);

            ArrayList temp = new ArrayList();
            while (rs.PeekNextClass() == typeof(SupBookRecord))
            {
                temp.Add(new ExternalBookBlock(rs));
            }

            //_externalBookBlocks = new ExternalBookBlock[temp.Count];
            _externalBookBlocks = (ExternalBookBlock[])temp.ToArray(typeof(ExternalBookBlock));
            temp.Clear();

            if (_externalBookBlocks.Length > 0)
            {
                // If any ExternalBookBlock present, there is always 1 of ExternSheetRecord
                if (rs.PeekNextClass() != typeof(ExternSheetRecord))
                {
                    // not quite - if written by google docs
                    _externSheetRecord = null;
                }
                else
                {
                    _externSheetRecord = ReadExtSheetRecord(rs);
                }
            }
            else
            {
                _externSheetRecord = null;
            }

            _definedNames = new List<NameRecord>();
            // collect zero or more DEFINEDNAMEs id=0x18
            while (true)
            {

                Type nextClass = rs.PeekNextClass();
                if (nextClass == typeof(NameRecord))
                {
                    NameRecord nr = (NameRecord)rs.GetNext();
                    _definedNames.Add(nr);
                }
                else if (nextClass == typeof(NameCommentRecord))
                {
                    NameCommentRecord ncr = (NameCommentRecord)rs.GetNext();
                    //commentRecords.Add(ncr.NameText, ncr);
                    commentRecords[ncr.NameText] = ncr;
                }
                else
                {
                    break;
                }
            }

            _recordCount = rs.GetCountRead();
            for (int i = startIndex; i < startIndex + _recordCount; i++)
            {
                _workbookRecordList.Records.Add(inputList[i]);
            }

        }

        public LinkTable(int numberOfSheets, WorkbookRecordList workbookRecordList)
        {
            _workbookRecordList = workbookRecordList;
            _definedNames = new List<NameRecord>();
            _externalBookBlocks = new ExternalBookBlock[] {
                new ExternalBookBlock(numberOfSheets),
            };
            _externSheetRecord = new ExternSheetRecord();
            _recordCount = 2;

            // tell _workbookRecordList about the 2 new records

            SupBookRecord supbook = _externalBookBlocks[0].GetExternalBookRecord();

            int idx = FindFirstRecordLocBySid(CountryRecord.sid);
            if (idx < 0)
            {
                throw new Exception("CountryRecord not found");
            }
            _workbookRecordList.Add(idx + 1, _externSheetRecord);
            _workbookRecordList.Add(idx + 1, supbook);
        }

        /**
         * TODO - would not be required if calling code used RecordStream or similar
         */
        public int RecordCount
        {
            get { return _recordCount; }
        }


        public NameRecord GetSpecificBuiltinRecord(byte builtInCode, int sheetNumber)
        {

            IEnumerator<NameRecord> iterator = _definedNames.GetEnumerator();
            while (iterator.MoveNext())
            {
                NameRecord record = iterator.Current;

                //print areas are one based
                if (record.BuiltInName == builtInCode && record.SheetNumber == sheetNumber)
                {
                    return record;
                }
            }

            return null;
        }

        public void RemoveBuiltinRecord(byte name, int sheetIndex)
        {
            //the name array is smaller so searching through it should be faster than
            //using the FindFirstXXXX methods
            NameRecord record = GetSpecificBuiltinRecord(name, sheetIndex);
            if (record != null)
            {
                _definedNames.Remove(record);
            }
            // TODO - do we need "Workbook.records.Remove(...);" similar to that in Workbook.RemoveName(int namenum) {}?
        }
        /**
         * @param extRefIndex as from a {@link Ref3DPtg} or {@link Area3DPtg}
         * @return -1 if the reference is to an external book
         */
        public int GetFirstInternalSheetIndexForExtIndex(int extRefIndex)
        {
            if (extRefIndex >= _externSheetRecord.NumOfRefs || extRefIndex < 0)
            {
                return -1;
            }
            return _externSheetRecord.GetFirstSheetIndexFromRefIndex(extRefIndex);
        }
        /**
         * @param extRefIndex as from a {@link Ref3DPtg} or {@link Area3DPtg}
         * @return -1 if the reference is to an external book
         */
        public int GetLastInternalSheetIndexForExtIndex(int extRefIndex)
        {
            if (extRefIndex >= _externSheetRecord.NumOfRefs || extRefIndex < 0)
            {
                return -1;
            }
            return _externSheetRecord.GetLastSheetIndexFromRefIndex(extRefIndex);
        }
        [Obsolete]
        public void UpdateIndexToInternalSheet(int extRefIndex, int offset)
        {
            _externSheetRecord.AdjustIndex(extRefIndex, offset);
        }

        public void RemoveSheet(int sheetIdx)
        {
            _externSheetRecord.RemoveSheet(sheetIdx);
        }

        public int NumNames
        {
            get
            {
                return _definedNames.Count;
            }
        }

        private int ExtendExternalBookBlocks(ExternalBookBlock newBlock)
        {
            ExternalBookBlock[] tmp = new ExternalBookBlock[_externalBookBlocks.Length + 1];
            Array.Copy(_externalBookBlocks, 0, tmp, 0, _externalBookBlocks.Length);
            tmp[tmp.Length - 1] = newBlock;
            _externalBookBlocks = tmp;

            return (_externalBookBlocks.Length - 1);
        }

        private int FindRefIndexFromExtBookIndex(int extBookIndex)
        {
            return _externSheetRecord.FindRefIndexFromExtBookIndex(extBookIndex);
        }
        /**
         * Finds the external name definition for the given name,
         *  optionally restricted by externsheet index, and returns
         *  (if found) as a NameXPtg.
         * @param sheetRefIndex The Extern Sheet Index to look for, or -1 if any
         */
        public NameXPtg GetNameXPtg(String name, int sheetRefIndex)
        {
            // first find any external book block that contains the name:
            for (int i = 0; i < _externalBookBlocks.Length; i++)
            {
                int definedNameIndex = _externalBookBlocks[i].GetIndexOfName(name);
                if (definedNameIndex < 0)
                {
                    continue;
                }
                // Found one
                int thisSheetRefIndex = FindRefIndexFromExtBookIndex(i);
                if (thisSheetRefIndex >= 0)
                {
                    // Check for the sheet index match, if requested
                    if (sheetRefIndex == -1 || thisSheetRefIndex == sheetRefIndex)
                    {
                        return new NameXPtg(thisSheetRefIndex, definedNameIndex);
                    }
                }
            }
            return null;
        }
        public NameRecord GetNameRecord(int index)
        {
            return (NameRecord)_definedNames[index];
        }

        public void AddName(NameRecord name)
        {
            _definedNames.Add(name);

            // TODO - this Is messy
            // Not the most efficient way but the other way was causing too many bugs
            int idx = FindFirstRecordLocBySid(ExternSheetRecord.sid);
            if (idx == -1) idx = FindFirstRecordLocBySid(SupBookRecord.sid);
            if (idx == -1) idx = FindFirstRecordLocBySid(CountryRecord.sid);
            int countNames = _definedNames.Count;
            _workbookRecordList.Add(idx + countNames, name);

        }
        /**
     * Register an external name in this workbook
     *
     * @param name  the name to register
     * @return a NameXPtg describing this name 
     */
        public NameXPtg AddNameXPtg(String name)
        {
            int extBlockIndex = -1;
            ExternalBookBlock extBlock = null;

            // find ExternalBlock for Add-In functions and remember its index
            for (int i = 0; i < _externalBookBlocks.Length; i++)
            {
                SupBookRecord ebr = _externalBookBlocks[i].GetExternalBookRecord();
                if (ebr.IsAddInFunctions)
                {
                    extBlock = _externalBookBlocks[i];
                    extBlockIndex = i;
                    break;
                }
            }
            // An ExternalBlock for Add-In functions was not found. Create a new one.
            if (extBlock == null)
            {
                extBlock = new ExternalBookBlock();

                extBlockIndex = ExtendExternalBookBlocks(extBlock);

                // add the created SupBookRecord before ExternSheetRecord
                int idx = FindFirstRecordLocBySid(ExternSheetRecord.sid);
                _workbookRecordList.Add(idx, extBlock.GetExternalBookRecord());

                // register the SupBookRecord in the ExternSheetRecord
                // -2 means that the scope of this name is Workbook and the reference applies to the entire workbook.
                _externSheetRecord.AddRef(_externalBookBlocks.Length - 1, -2, -2);
            }

            // create a ExternalNameRecord that will describe this name
            ExternalNameRecord extNameRecord = new ExternalNameRecord();
            extNameRecord.Text = (name);
            // The docs don't explain why Excel set the formula to #REF!
            extNameRecord.SetParsedExpression(new Ptg[] { ErrPtg.REF_INVALID });

            int nameIndex = extBlock.AddExternalName(extNameRecord);
            int supLinkIndex = 0;
            // find the posistion of the Add-In SupBookRecord in the workbook stream,
            // the created ExternalNameRecord will be appended to it
            for (IEnumerator iterator = _workbookRecordList.GetEnumerator(); iterator.MoveNext(); supLinkIndex++)
            {
                Record record = (Record)iterator.Current;
                if (record is SupBookRecord)
                {
                    if (((SupBookRecord)record).IsAddInFunctions) break;
                }
            }
            int numberOfNames = extBlock.NumberOfNames;
            // a new name is inserted in the end of the SupBookRecord, after the last name
            _workbookRecordList.Add(supLinkIndex + numberOfNames, extNameRecord);
            int fakeSheetIdx = -2; /* the scope is workbook*/
            int ix = _externSheetRecord.GetRefIxForSheet(extBlockIndex, fakeSheetIdx, fakeSheetIdx);
            return new NameXPtg(ix, nameIndex);
        }
        public void RemoveName(int namenum)
        {
            _definedNames.RemoveAt(namenum);
        }


        private static int GetSheetIndex(String[] sheetNames, String sheetName)
        {
            for (int i = 0; i < sheetNames.Length; i++)
            {
                if (sheetNames[i].Equals(sheetName))
                {
                    return i;
                }

            }
            throw new InvalidOperationException("External workbook does not contain sheet '" + sheetName + "'");
        }
        private int GetExternalWorkbookIndex(String workbookName)
        {
            for (int i = 0; i < _externalBookBlocks.Length; i++)
            {
                SupBookRecord ebr = _externalBookBlocks[i].GetExternalBookRecord();
                if (!ebr.IsExternalReferences)
                {
                    continue;
                }
                if (workbookName.Equals(ebr.URL))
                { // not sure if 'equals()' works when url has a directory
                    return i;
                }
            }
            return -1;
        }

        public int LinkExternalWorkbook(String name, IWorkbook externalWorkbook)
        {
            int extBookIndex = GetExternalWorkbookIndex(name);
            if (extBookIndex != -1)
            {
                // Already linked!
                return extBookIndex;
            }

            // Create a new SupBookRecord
            String[] sheetNames = new String[externalWorkbook.NumberOfSheets];
            for (int sn = 0; sn < sheetNames.Length; sn++)
            {
                sheetNames[sn] = externalWorkbook.GetSheetName(sn);
            }
            //\000 is octal digit in java, but c# think it is a '\0' and two zero.
            String url = "\0" + name;
            ExternalBookBlock block = new ExternalBookBlock(url, sheetNames);

            // Add it into the list + records
            extBookIndex = ExtendExternalBookBlocks(block);

            // add the created SupBookRecord before ExternSheetRecord
            int idx = FindFirstRecordLocBySid(ExternSheetRecord.sid);
            if (idx == -1)
            {
                idx = _workbookRecordList.Count;
            }
            _workbookRecordList.Add(idx, block.GetExternalBookRecord());

            // Setup links for the sheets
            for (int sn = 0; sn < sheetNames.Length; sn++)
            {
                _externSheetRecord.AddRef(extBookIndex, sn, sn);
            }

            // Report where it went
            return extBookIndex;
        }

        public int GetExternalSheetIndex(String workbookName, String firstSheetName, String lastSheetName)
        {
            int externalBookIndex = GetExternalWorkbookIndex(workbookName);
            if (externalBookIndex == -1)
            {
                throw new RuntimeException("No external workbook with name '" + workbookName + "'");
            }
            SupBookRecord ebrTarget = _externalBookBlocks[externalBookIndex].GetExternalBookRecord();

            int firstSheetIndex = GetSheetIndex(ebrTarget.SheetNames, firstSheetName);
            int lastSheetIndex = GetSheetIndex(ebrTarget.SheetNames, lastSheetName);

            // Find or add the external sheet record definition for this
            int result = _externSheetRecord.GetRefIxForSheet(externalBookIndex, firstSheetIndex, lastSheetIndex);
            if (result < 0)
            {
                result = _externSheetRecord.AddRef(externalBookIndex, firstSheetIndex, lastSheetIndex);
            }
            return result;
        }
        public String[] GetExternalBookAndSheetName(int extRefIndex)
        {
            int ebIx = _externSheetRecord.GetExtbookIndexFromRefIndex(extRefIndex);
            SupBookRecord ebr = _externalBookBlocks[ebIx].GetExternalBookRecord();
            if (!ebr.IsExternalReferences)
            {
                return null;
            }
            // Sheet name only applies if not a global reference
            int shIx1 = _externSheetRecord.GetFirstSheetIndexFromRefIndex(extRefIndex);
            int shIx2 = _externSheetRecord.GetLastSheetIndexFromRefIndex(extRefIndex);
            String firstSheetName = null;
            String lastSheetName = null;
            if (shIx1 >= 0)
            {
                firstSheetName = ebr.SheetNames[shIx1];
            }
            if (shIx2 >= 0)
            {
                lastSheetName = ebr.SheetNames[shIx2];
            }
            if (shIx1 == shIx2)
            {
                return new String[] {
    				ebr.URL,
    				firstSheetName
    		};
            }
            else
            {
                return new String[] {
                    ebr.URL,
                    firstSheetName,
                    lastSheetName
            };
            }
        }
        public int CheckExternSheet(int sheetIndex)
        {
            return CheckExternSheet(sheetIndex, sheetIndex);
        }
        public int CheckExternSheet(int firstSheetIndex, int lastSheetIndex)
        {
            int thisWbIndex = -1; // this is probably always zero
            for (int i = 0; i < _externalBookBlocks.Length; i++)
            {
                SupBookRecord ebr = _externalBookBlocks[i].GetExternalBookRecord();
                if (ebr.IsInternalReferences)
                {
                    thisWbIndex = i;
                    break;
                }
            }
            if (thisWbIndex < 0)
            {
                throw new InvalidOperationException("Could not find 'internal references' EXTERNALBOOK");
            }

            //Trying to find reference to this sheet
            int j = _externSheetRecord.GetRefIxForSheet(thisWbIndex, firstSheetIndex, lastSheetIndex);
            if (j >= 0)
            {
                return j;
            }
            //We haven't found reference to this sheet
            return _externSheetRecord.AddRef(thisWbIndex, firstSheetIndex, lastSheetIndex);

        }


        /**
         * copied from Workbook
         */
        private int FindFirstRecordLocBySid(short sid)
        {
            int index = 0;
            for (IEnumerator<Record> iterator = _workbookRecordList.GetEnumerator(); iterator.MoveNext(); )
            {
                Record record = iterator.Current;

                if (record.Sid == sid)
                {
                    return index;
                }
                index++;
            }
            return -1;
        }

        public String ResolveNameXText(int refIndex, int definedNameIndex, InternalWorkbook workbook)
        {
            int extBookIndex = _externSheetRecord.GetExtbookIndexFromRefIndex(refIndex);
            int firstTabIndex = _externSheetRecord.GetFirstSheetIndexFromRefIndex(refIndex);
            if (firstTabIndex == -1)
            {
                // The referenced sheet could not be found
                throw new RuntimeException("Referenced sheet could not be found");
            }

            // Does it exist via the external book block?
            ExternalBookBlock externalBook = _externalBookBlocks[extBookIndex];
            if (externalBook._externalNameRecords.Length > definedNameIndex)
            {
                return _externalBookBlocks[extBookIndex].GetNameText(definedNameIndex);
            }
            else if (firstTabIndex == -2)
            {
                // Workbook scoped name, not actually external after all
                NameRecord nr = GetNameRecord(definedNameIndex);
                int sheetNumber = nr.SheetNumber;

                StringBuilder text = new StringBuilder();
                if (sheetNumber > 0)
                {
                    String sheetName = workbook.GetSheetName(sheetNumber - 1);
                    SheetNameFormatter.AppendFormat(text, sheetName);
                    text.Append("!");
                }
                text.Append(nr.NameText);
                return text.ToString();
            }
            else
            {
                throw new IndexOutOfRangeException(
                        "Ext Book Index relative but beyond the supported length, was " +
                        extBookIndex + " but maximum is " + _externalBookBlocks.Length
                );
            }
        }
        public int ResolveNameXIx(int refIndex, int definedNameIndex)
        {
            int extBookIndex = _externSheetRecord.GetExtbookIndexFromRefIndex(refIndex);
            return _externalBookBlocks[extBookIndex].GetNameIx(definedNameIndex);
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
            foreach (ExternalBookBlock ex in _externalBookBlocks)
            {
                SupBookRecord externalRecord = ex.GetExternalBookRecord();
                if (externalRecord.IsExternalReferences
                    && externalRecord.URL.Equals(oldUrl))
                {
                    externalRecord.URL = (newUrl);
                    return true;
                }
            }
            return false;
        }
    }
}
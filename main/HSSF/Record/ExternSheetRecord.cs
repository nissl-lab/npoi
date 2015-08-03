
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


namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using System.Collections;
    using NPOI.Util;
    using System.Collections.Generic;


    public class RefSubRecord
    {
        public const int ENCODED_SIZE = 6;

        /** index to External Book Block (which starts with a EXTERNALBOOK record) */
        private int _extBookIndex;
        private int _firstSheetIndex; // may be -1 (0xFFFF)
        private int _lastSheetIndex;  // may be -1 (0xFFFF)

        public void AdjustIndex(int offset)
        {
            _firstSheetIndex += offset;
            _lastSheetIndex += offset;
        }
        /** a Constructor for making new sub record
         */
        public RefSubRecord(int extBookIndex, int firstSheetIndex, int lastSheetIndex)
        {
            _extBookIndex = extBookIndex;
            _firstSheetIndex = firstSheetIndex;
            _lastSheetIndex = lastSheetIndex;
        }

        /**
         * @param in the RecordInputstream to Read the record from
         */
        public RefSubRecord(RecordInputStream in1)
            : this(in1.ReadShort(), in1.ReadShort(), in1.ReadShort())
        {

        }
        public int ExtBookIndex
        {
            get
            {
                return _extBookIndex;
            }
        }
        public int FirstSheetIndex
        {
            get
            {
                return _firstSheetIndex;
            }
        }
        public int LastSheetIndex
        {
            get { return _lastSheetIndex; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("extBook=").Append(_extBookIndex);
            buffer.Append(" firstSheet=").Append(_firstSheetIndex);
            buffer.Append(" lastSheet=").Append(_lastSheetIndex);
            return buffer.ToString();
        }

        /**
         * called by the class that is responsible for writing this sucker.
         * Subclasses should implement this so that their data is passed back in a
         * byte array.
         *
         * @param offset to begin writing at
         * @param data byte array containing instance data
         * @return number of bytes written
         */
        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_extBookIndex);
            out1.WriteShort(_firstSheetIndex);
            out1.WriteShort(_lastSheetIndex);
        }
    }

    /**
     * Title:        Extern Sheet 
     * Description:  A List of Inndexes to SupBook 
     * REFERENCE:  
     * @author Libin Roman (Vista Portal LDT. Developer)
     * @version 1.0-pre
     */

    public class ExternSheetRecord : StandardRecord
    {
        public const short sid = 0x17;
        private IList<RefSubRecord> _list;




        public ExternSheetRecord()
        {
            _list = new List<RefSubRecord>();
        }

        /**
         * Constructs a Extern Sheet record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public ExternSheetRecord(RecordInputStream in1)
        {
            _list = new List<RefSubRecord>();

            int nItems = in1.ReadShort();

            for (int i = 0; i < nItems; ++i)
            {
                RefSubRecord rec = new RefSubRecord(in1);

                _list.Add(rec);
            }
        }
        /**
        * Add a zero-based reference to a {@link org.apache.poi.hssf.record.SupBookRecord}.
        * <p>
        *  If the type of the SupBook record is same-sheet referencing, Add-In referencing,
        *  DDE data source referencing, or OLE data source referencing,
        *  then no scope is specified and this value <em>MUST</em> be -2. Otherwise,
        *  the scope must be set as follows:
        *   <ol>
        *    <li><code>-2</code> Workbook-level reference that applies to the entire workbook.</li>
        *    <li><code>-1</code> Sheet-level reference. </li>
        *    <li><code>&gt;=0</code> Sheet-level reference. This specifies the first sheet in the reference.
        *    <p>
        *    If the SupBook type is unused or external workbook referencing,
        *    then this value specifies the zero-based index of an external sheet name,
        *    see {@link org.apache.poi.hssf.record.SupBookRecord#getSheetNames()}.
        *    This referenced string specifies the name of the first sheet within the external workbook that is in scope.
        *    This sheet MUST be a worksheet or macro sheet.
        *    </p>
        *    <p>
        *    If the supporting link type is self-referencing, then this value specifies the zero-based index of a
        *    {@link org.apache.poi.hssf.record.BoundSheetRecord} record in the workbook stream that specifies
        *    the first sheet within the scope of this reference. This sheet MUST be a worksheet or a macro sheet.
        *    </p>
        *    </li>
        *  </ol></p>
        *
        * @param firstSheetIndex  the scope, must be -2 for add-in references
        * @param lastSheetIndex   the scope, must be -2 for add-in references
        * @return index of newly added ref
        */
        public int AddRef(int extBookIndex, int firstSheetIndex, int lastSheetIndex)
        {
            _list.Add(new RefSubRecord(extBookIndex, firstSheetIndex, lastSheetIndex));
            return _list.Count - 1;
        }
        public int GetRefIxForSheet(int externalBookIndex, int firstSheetIndex, int lastSheetIndex)
        {
            int nItems = _list.Count;
            for (int i = 0; i < nItems; i++)
            {
                RefSubRecord ref1 = GetRef(i);
                if (ref1.ExtBookIndex != externalBookIndex)
                {
                    continue;
                }
                if (ref1.FirstSheetIndex == firstSheetIndex && ref1.LastSheetIndex == lastSheetIndex)
                {
                    return i;
                }
            }
            return -1;
        }
        /** returns the number of REF Records, which is in model
         * @return number of REF records
         */
        public int NumOfREFRecords
        {
            get
            {
                return _list.Count;
            }
        }
    
        /**  
 * @return number of REF structures
 */
        public int NumOfRefs
        {
            get
            {
                return _list.Count;
            }
        }

        /** 
         * Adds REF struct (ExternSheetSubRecord)
         * @param rec REF struct
         */
        public void AddREFRecord(RefSubRecord rec)
        {
            _list.Add(rec);
        }
        private RefSubRecord GetRef(int i)
        {
            return (RefSubRecord)_list[i];
        }
        [Obsolete]
        public void AdjustIndex(int extRefIndex, int offset)
        {
            GetRef(extRefIndex).AdjustIndex(offset);
        }

        public void RemoveSheet(int sheetIdx)
        {
            int nItems = _list.Count;
            int toRemove = -1;
            for (int i = 0; i < nItems; i++)
            {
                RefSubRecord refSubRecord = _list[(i)];
                if (refSubRecord.FirstSheetIndex == sheetIdx &&
                        refSubRecord.LastSheetIndex == sheetIdx)
                {
                    toRemove = i;
                }
                else if (refSubRecord.FirstSheetIndex > sheetIdx &&
                      refSubRecord.LastSheetIndex > sheetIdx)
                {
                    _list[i] =(new RefSubRecord(refSubRecord.ExtBookIndex, refSubRecord.FirstSheetIndex - 1, refSubRecord.LastSheetIndex - 1));
                }
            }

            // finally remove entries for sheet indexes that we remove
            if (toRemove != -1)
            {
                _list.RemoveAt(toRemove);
            }
        }

        /**
         * Returns the index of the SupBookRecord for this index
         */
        public int GetExtbookIndexFromRefIndex(int refIndex)
        {
            RefSubRecord refRec = GetRef(refIndex);
            return refRec.ExtBookIndex;
        }
        /**
         * @return -1 if not found
         */
        public int FindRefIndexFromExtBookIndex(int extBookIndex)
        {
            int nItems = _list.Count;
            for (int i = 0; i < nItems; i++)
            {
                if (GetRef(i).ExtBookIndex == extBookIndex)
                {
                    return i;
                }
            }
            return -1;
        }
        public static ExternSheetRecord Combine(ExternSheetRecord[] esrs)
        {
            ExternSheetRecord result = new ExternSheetRecord();
            for (int i = 0; i < esrs.Length; i++)
            {
                ExternSheetRecord esr = esrs[i];
                int nRefs = esr.NumOfREFRecords;
                for (int j = 0; j < nRefs; j++)
                {
                    result.AddREFRecord(esr.GetRef(j));
                }
            }
            return result;
        }
        /**
         * Returns the first sheet that the reference applies to, or
         *  -1 if the referenced sheet can't be found, or -2 if the
         *  reference is workbook scoped.
         */
        public int GetFirstSheetIndexFromRefIndex(int extRefIndex)
        {
            return GetRef(extRefIndex).FirstSheetIndex;
        }

        /**
         * Returns the last sheet that the reference applies to, or
         *  -1 if the referenced sheet can't be found, or -2 if the
         *  reference is workbook scoped.
         * For a single sheet reference, the first and last should be
         *  the same.
         */
        public int GetLastSheetIndexFromRefIndex(int extRefIndex)
        {
            return GetRef(extRefIndex).LastSheetIndex;
        }
    
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            int nItems = _list.Count;
            sb.Append("[EXTERNSHEET]\n");
            sb.Append("   numOfRefs     = ").Append(nItems).Append("\n");
            for (int i = 0; i < nItems; i++)
            {
                sb.Append("refrec         #").Append(i).Append(": ");
                sb.Append(GetRef(i).ToString());
                sb.Append('\n');
            }
            sb.Append("[/EXTERNSHEET]\n");

            return sb.ToString();
        }

        /**
         * called by the class that Is responsible for writing this sucker.
         * Subclasses should implement this so that their data Is passed back in a
         * byte array.
         *
         * @param offset to begin writing at
         * @param data byte array containing instance data
         * @return number of bytes written
         */
        public override void Serialize(ILittleEndianOutput out1)
        {
            int nItems = _list.Count;

            out1.WriteShort(nItems);

            for (int i = 0; i < nItems; i++)
            {
                GetRef(i).Serialize(out1);
            }
        }

        protected override int DataSize
        {
            get { 
                return 2 + _list.Count * RefSubRecord.ENCODED_SIZE; 
            }
        }

        /**
         * return the non static version of the id for this record.
         */
        public override short Sid
        {
            get { return sid; }
        }
    }
}
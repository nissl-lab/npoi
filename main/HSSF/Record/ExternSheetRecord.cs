
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


    public class RefSubRecord
    {
        public const int ENCODED_SIZE = 6;

        /** index to External Book Block (which starts with a EXTERNALBOOK record) */
        private int _extBookIndex;
        private int _firstSheetIndex; // may be -1 (0xFFFF)
        private int _lastSheetIndex;  // may be -1 (0xFFFF)


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
        private IList _list;




        public ExternSheetRecord()
        {
            _list = new ArrayList();
        }

        /**
         * Constructs a Extern Sheet record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public ExternSheetRecord(RecordInputStream in1)
        {
            _list = new ArrayList();

            int nItems = in1.ReadShort();

            for (int i = 0; i < nItems; ++i)
            {
                RefSubRecord rec = new RefSubRecord(in1);

                _list.Add(rec);
            }
        }
        /**
 * @return index of newly added ref
 */
        public int AddRef(int extBookIndex, int firstSheetIndex, int lastSheetIndex)
        {
            _list.Add(new RefSubRecord(extBookIndex, firstSheetIndex, lastSheetIndex));
            return _list.Count - 1;
        }
        public int GetRefIxForSheet(int externalBookIndex, int sheetIndex)
        {
            int nItems = _list.Count;
            for (int i = 0; i < nItems; i++)
            {
                RefSubRecord ref1 = GetRef(i);
                if (ref1.ExtBookIndex != externalBookIndex)
                {
                    continue;
                }
                if (ref1.FirstSheetIndex == sheetIndex && ref1.LastSheetIndex == sheetIndex)
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
        public int GetExtbookIndexFromRefIndex(int refIndex)
        {
            return GetRef(refIndex).ExtBookIndex;
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
        public int GetFirstSheetIndexFromRefIndex(int extRefIndex)
        {
            return GetRef(extRefIndex).FirstSheetIndex;
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
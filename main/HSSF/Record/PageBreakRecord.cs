
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
    using System.Collections;
    using System.Text;
    using NPOI.Util;
    using System.Collections.Generic;


    /**
     * Record that Contains the functionality page _breaks (horizontal and vertical)
     * 
     * The other two classes just specifically Set the SIDS for record creation.
     * 
     * REFERENCE:  Microsoft Excel SDK page 322 and 420
     * 
     * @see HorizontalPageBreakRecord
     * @see VerticalPageBreakRecord
     * @author Danny Mui (dmui at apache dot org)
     */
    public class PageBreakRecord : StandardRecord
    {
        private const bool IS_EMPTY_RECORD_WRITTEN = false;
        private static readonly int[] EMPTY_INT_ARRAY = { };

        public short sid;
        // fix warning CS0169 "never used": private short numBreaks;
        private IList<Break> _breaks;
        private Hashtable _breakMap;

        /**
         * Since both records store 2byte integers (short), no point in 
         * differentiating it in the records.
         * 
         * The subs (rows or columns, don't seem to be able to Set but excel Sets
         * them automatically)
         */
        public class Break
        {

            public const int ENCODED_SIZE = 6;
            public int main;
            public int subFrom;
            public int subTo;

            public Break(RecordInputStream in1)
            {
                main = in1.ReadUShort() - 1;
                subFrom = in1.ReadUShort();
                subTo = in1.ReadUShort();
            }

            public Break(int main, int subFrom, int subTo)
            {
                this.main = main;
                this.subFrom = subFrom;
                this.subTo = subTo;
            }

            public void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(main + 1);
                out1.WriteShort(subFrom);
                out1.WriteShort(subTo);
            }
        }

        public PageBreakRecord()
        {
            _breaks = new List<Break>();
            _breakMap = new Hashtable();
        }

        public PageBreakRecord(RecordInputStream in1)
        {
            int nBreaks = in1.ReadShort();
            _breaks = new List<Break>(nBreaks + 2);
            _breakMap = new Hashtable();

            for (int k = 0; k < nBreaks; k++)
            {
                Break br = new Break(in1);
                _breaks.Add(br);
                _breakMap[br.main] = br;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        //public override int Serialize(int offset, byte[] data)
        //{
        //    int nBreaks = _breaks.Count;
        //    if (!IS_EMPTY_RECORD_WRITTEN && nBreaks < 1)
        //    {
        //        return 0;
        //    }
        //    int dataSize = DataSize;
        //    LittleEndian.PutUShort(data, offset + 0, Sid);
        //    LittleEndian.PutUShort(data, offset + 2, dataSize);
        //    LittleEndian.PutUShort(data, offset + 4, nBreaks);
        //    int pos = 6;
        //    for (int i = 0; i < nBreaks; i++)
        //    {
        //        Break br = (Break)_breaks[i];
        //        pos += br.Serialize(offset + pos, data);
        //    }

        //    return 4 + dataSize;
        //}
        public override void Serialize(ILittleEndianOutput out1)
        {
            int nBreaks = _breaks.Count;
            out1.WriteShort(nBreaks);
            for (int i = 0; i < nBreaks; i++)
            {
                _breaks[i].Serialize(out1);
            }
        }
        protected override int DataSize
        {
            get
            {
                return 2 + _breaks.Count * Break.ENCODED_SIZE;
            }
        }

        public IEnumerator<Break> GetBreaksEnumerator()
        {
            //if (_breaks == null)
            //    return new ArrayList().GetEnumerator();
            //else
            return _breaks.GetEnumerator();
        }

        public override String ToString()
        {
            StringBuilder retval = new StringBuilder();

            //if (Sid != HORIZONTAL_SID && Sid != VERTICAL_SID)
            //    return "[INVALIDPAGEBREAK]\n     .Sid =" + Sid + "[INVALIDPAGEBREAK]";

            String label;
            String mainLabel;
            String subLabel;

            if (Sid == HorizontalPageBreakRecord.sid)
            {
                label = "HORIZONTALPAGEBREAK";
                mainLabel = "row";
                subLabel = "col";
            }
            else
            {
                label = "VERTICALPAGEBREAK";
                mainLabel = "column";
                subLabel = "row";
            }

            retval.Append("[" + label + "]").Append("\n");
            retval.Append("     .Sid        =").Append(Sid).Append("\n");
            retval.Append("     .num_breaks =").Append(NumBreaks).Append("\n");
            IEnumerator iterator = GetBreaksEnumerator();
            for (int k = 0; k < NumBreaks; k++)
            {
                Break region = (Break)iterator.Current;

                retval.Append("     .").Append(mainLabel).Append(" (zero-based) =").Append(region.main).Append("\n");
                retval.Append("     .").Append(subLabel).Append("From    =").Append(region.subFrom).Append("\n");
                retval.Append("     .").Append(subLabel).Append("To      =").Append(region.subTo).Append("\n");
            }

            retval.Append("[" + label + "]").Append("\n");
            return retval.ToString();
        }

        /**
         * Adds the page break at the specified parameters
         * @param main Depending on sid, will determine row or column to put page break (zero-based)
         * @param subFrom No user-interface to Set (defaults to minumum, 0)
         * @param subTo No user-interface to Set
         */
        public void AddBreak(int main, int subFrom, int subTo)
        {
            //if (_breaks == null)
            //{
            //    _breaks = new ArrayList(NumBreaks + 10);
            //    _breakMap = new Hashtable();
            //}
            int key = (int)main;
            Break region = (Break)_breakMap[key];
            if (region != null)
            {
                region.main = main;
                region.subFrom = subFrom;
                region.subTo = subTo;
            }
            else
            {
                region = new Break(main, subFrom, subTo);
                _breaks.Add(region);
            }
            _breakMap[key] = region;
        }

        /**
         * Removes the break indicated by the parameter
         * @param main (zero-based)
         */
        public void RemoveBreak(int main)
        {
            int rowKey = main;
            Break region = (Break)_breakMap[rowKey];
            _breaks.Remove(region);
            _breakMap.Remove(rowKey);
        }

        public override int RecordSize
        {
            get {
                int nBreaks = _breaks.Count;
                if (!IS_EMPTY_RECORD_WRITTEN && nBreaks < 1)
                {
                    return 0;
                }
                return 4 + DataSize;
            }
        }
        public int NumBreaks
        {
            get
            {
                return _breaks.Count;
            }
        }
        public bool IsEmpty
        {
            get
            {
                return _breaks.Count==0;
            }
        }
        /**
         * Retrieves the region at the row/column indicated
         * @param main FIXME: Document this!
         * @return The Break or null if no break exists at the row/col specified.
         */
        public Break GetBreak(int main)
        {
            //if (_breakMap == null)
            //    return null;
            //int rowKey = (int)main;
            return (Break)_breakMap[main];
        }
        public int[] GetBreaks()
        {
            int count = NumBreaks;
            if (count < 1)
            {
                return EMPTY_INT_ARRAY;
            }
            int[] result = new int[count];
            for (int i = 0; i < count; i++)
            {
                Break breakItem = _breaks[i];
                result[i] = breakItem.main;
            }
            return result;
        }
    }
}

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

namespace NPOI.HSSF.Record
{
    using NPOI.SS.Util;
    using NPOI.Util;


    /**
     * Common base class for {@link SharedFormulaRecord}, {@link ArrayRecord} and
     * {@link TableRecord} which are have similarities.
     * 
     * @author Josh Micich
     */
    public abstract class SharedValueRecordBase : StandardRecord
    {

        private CellRangeAddress8Bit _range;

        protected SharedValueRecordBase(CellRangeAddress8Bit range)
        {
            _range = range;
        }

        protected SharedValueRecordBase()
            : this(new CellRangeAddress8Bit(0, 0, 0, 0))
        {

        }

        /**
         * reads only the range (1 {@link CellRangeAddress8Bit}) from the stream
         */
        public SharedValueRecordBase(RecordInputStream in1)
        {
            _range = new CellRangeAddress8Bit(in1);
        }

        public CellRangeAddress8Bit Range
        {
            get
            {
                return _range;
            }
        }

        public virtual int FirstRow
        {
            get
            {
                return _range.FirstRow;
            }
        }

        public virtual int LastRow
        {
            get
            {
                return _range.LastRow;
            }
        }

        public virtual int FirstColumn
        {
            get
            {
                return (short)_range.FirstColumn;
            }
        }

        public virtual int LastColumn
        {
            get
            {
                return (short)_range.LastColumn;
            }
        }
        protected override int DataSize
        {
            get
            {
                return CellRangeAddress8Bit.ENCODED_SIZE + this.ExtraDataSize;
            }
        }
        protected abstract int ExtraDataSize { get; }

        protected abstract void SerializeExtraData(ILittleEndianOutput out1);

        public override void Serialize(ILittleEndianOutput out1)
        {
            _range.Serialize(out1);
            SerializeExtraData(out1);
        }

        /**
         * @return <c>true</c> if (rowIx, colIx) is within the range ({@link #Range})
         * of this shared value object.
         */
        public bool IsInRange(int rowIx, int colIx)
        {
            CellRangeAddress8Bit r = _range;
            return r.FirstRow <= rowIx
                && r.LastRow >= rowIx
                && r.FirstColumn <= colIx
                && r.LastColumn >= colIx;
        }
        /**
         * @return <c>true</c> if (rowIx, colIx) describes the first cell in this shared value 
         * object's range ({@link #Range})
         */
        public bool IsFirstCell(int rowIx, int colIx)
        {
            CellRangeAddress8Bit r = Range;
            return r.FirstRow == rowIx && r.FirstColumn == colIx;
        }
    }
}
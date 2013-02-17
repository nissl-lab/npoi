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

namespace NPOI.SS.Formula.PTG
{
    using System;
    using System.Text;
    using NPOI.Util;

    using NPOI.SS.Util;

    /**
     * @author Josh Micich
     */
    [Serializable]
    public abstract class Ref2DPtgBase : RefPtgBase
    {
        private const int SIZE = 5;

        /**
         * Takes in a String representation of a cell reference and fills out the
         * numeric fields.
         */
        protected Ref2DPtgBase(String cellref)
            : base(cellref)
        {

        }
        protected Ref2DPtgBase(CellReference cr):base(cr)
        {
            
        }

        protected Ref2DPtgBase(int row, int column, bool isRowRelative, bool isColumnRelative)
        {
            Row = (row);
            Column = (column);
            IsRowRelative = (isRowRelative);
            IsColRelative = (isColumnRelative);
        }

        protected Ref2DPtgBase(ILittleEndianInput in1)
        {
            ReadCoordinates(in1);
        }
        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(Sid + PtgClass);
            WriteCoordinates(out1);
        }
        public override String ToFormulaString()
        {
            return FormatReferenceAsString();
        }

        protected abstract byte Sid { get; }

        public override int Size
        {
            get
            {
                return SIZE;
            }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name);
            sb.Append(" [");
            sb.Append(FormatReferenceAsString());
            sb.Append("]");
            return sb.ToString();
        }
    }
}
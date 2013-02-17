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
     * Common superclass of 2-D area refs 
     */
    [Serializable]
    public abstract class Area2DPtgBase : AreaPtgBase
    {
        private const int SIZE = 9;

        protected Area2DPtgBase(int firstRow, int lastRow, int firstColumn, int lastColumn, bool  firstRowRelative, bool  lastRowRelative, bool  firstColRelative, bool  lastColRelative)
            : base(firstRow, lastRow, firstColumn, lastColumn, firstRowRelative, lastRowRelative, firstColRelative, lastColRelative)
        {

        }
        protected Area2DPtgBase(AreaReference ar):base(ar)
        {
            
        }
        protected Area2DPtgBase(ILittleEndianInput in1)
        {
            ReadCoordinates(in1);
        }
        protected abstract byte Sid { get; }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(Sid + PtgClass);
            WriteCoordinates(out1);
        }
        public Area2DPtgBase(String arearef)
            : base(arearef)
        {
        }
        public override int Size
        {
            get
            {
                return SIZE;
            }
        }
        public override String ToFormulaString()
        {
            return FormatReferenceAsString();
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
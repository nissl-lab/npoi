/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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
    using System;

    using NPOI.HSSF.Record.Common;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * Conditional Formatting Header v12 record CFHEADER12 (0x0879),
     *  for conditional formattings introduced in Excel 2007 and newer.
     */
    public class CFHeader12Record : CFHeaderBase, IFutureRecord, ICloneable
    {
        public static short sid = 0x0879;

        private FtrHeader futureHeader;

        /** Creates new CFHeaderRecord */
        public CFHeader12Record()
        {
            CreateEmpty();
            futureHeader = new FtrHeader();
            futureHeader.RecordType = (/*setter*/sid);
        }
        public CFHeader12Record(CellRangeAddress[] regions, int nRules)
                : base(regions, nRules)
        {
            ;
            futureHeader = new FtrHeader();
            futureHeader.RecordType = (/*setter*/sid);
        }
        public CFHeader12Record(RecordInputStream in1)
        {
            futureHeader = new FtrHeader(in1);
            Read(in1);
        }


        protected override string RecordName
        {
            get
            {
                return "CFHEADER12";
            }
        }

        protected override int DataSize
        {
            get
            {
                return FtrHeader.GetDataSize() + base.DataSize;
            }

        }
        public override void Serialize(ILittleEndianOutput out1)
        {
            // Sync the associated range
            futureHeader.AssociatedRange = (EnclosingCellRange);
            // Write the future header first
            futureHeader.Serialize(out1);
            // Then the rest of the CF Header details
            base.Serialize(out1);
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public short GetFutureRecordType()
        {
            return futureHeader.RecordType;
        }
        public FtrHeader GetFutureHeader()
        {
            return futureHeader;
        }
        public CellRangeAddress GetAssociatedRange()
        {
            return futureHeader.AssociatedRange;
        }

        public override object Clone()
        {
            CFHeader12Record result = new CFHeader12Record();
            result.futureHeader = (FtrHeader)futureHeader.Clone();
            base.CopyTo(result);
            return result;
        }
    }
}

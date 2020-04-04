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

namespace NPOI.HSSF.Record.Common
{
    using System;

    using NPOI.HSSF.Record;
    using NPOI.Util;
    using System.Text;
    using NPOI.SS.Util;

    /**
     * Title: FtrHeader (Future Record Header) common record part
     * 
     * This record part specifies a header for a Ftr (Future)
     *  style record, which includes extra attributes above and
     *  beyond those of a traditional record. 
     */
    public class FtrHeader : ICloneable
    {
        /** This MUST match the type on the Containing record */
        private short recordType;
        /** This is a FrtFlags */
        private short grbitFrt;

        /** The range of cells the parent record applies to, or 0 if N/A */
        private CellRangeAddress associatedRange;
        public FtrHeader()
        {
            associatedRange = new CellRangeAddress(0, 0, 0, 0);
        }

        public FtrHeader(RecordInputStream in1)
        {
            recordType = in1.ReadShort();
            grbitFrt = in1.ReadShort();

            associatedRange = new CellRangeAddress(in1);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append(" [FUTURE HEADER]\n");
            buffer.Append("   type " + recordType);
            buffer.Append("   flags " + grbitFrt);
            buffer.Append(" [/FUTURE HEADER]\n");
            return buffer.ToString();
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(recordType);
            out1.WriteShort(grbitFrt);
            associatedRange.Serialize(out1);
        }

        public static int GetDataSize()
        {
            return 12;
        }

        public short RecordType
        {
            get
            {
                return recordType;
            }
            set
            {
                recordType = value;
            }
        }

        public short GrbitFrt
        {
            get
            {
                return grbitFrt;
            }
            set
            {
                grbitFrt = value;
            }
        }

 
        public CellRangeAddress AssociatedRange
        {
            get
            {
                return associatedRange;
            }
            set
            {
                this.associatedRange = value;
            }
        }

        public object Clone()
        {
            FtrHeader result = new FtrHeader
            {
                recordType = recordType,
                grbitFrt = grbitFrt,
                associatedRange = associatedRange.Copy()
            };
            return result;
        }
    }
}
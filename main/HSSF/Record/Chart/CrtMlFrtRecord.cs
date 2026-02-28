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

namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// The CrtMlFrt record specifies additional properties for chart elements, as specified by 
    /// the Chart Sheet Substream ABNF. These properties complement the record to which they 
    /// correspond, and are stored as a structure chain defined in XmlTkChain. An application 
    /// can ignore this record without loss of functionality, except for the additional properties. 
    /// If this record is longer than 8224 bytes, it MUST be split into several records. The first
    /// section of the data appears in this record and subsequent sections appear in one or more 
    /// CrtMlFrtContinue records that follow this record.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class CrtMlFrtRecord : RowDataRecord
    {
        public const short sid = 0x89E;
        public CrtMlFrtRecord(RecordInputStream ris)
            : base(ris)
        {
        }

        protected override int DataSize
        {
            get
            {
                return base.DataSize;
            }
        }
        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            base.Serialize(out1);
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}

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
    /// <summary>
    /// The ContinueFrt12 record specifies a continuation of the data in a preceding Future Record 
    /// Type record that has data longer than 8,224 bytes. Such records are split into several records. 
    /// The first section of the data appears in the base record and subsequent sections appear in 
    /// one or more ContinueFrt12 records that appear after the base record. The preceding base record 
    /// MUST contain a FrtRefHeader or a FrtHeader field.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class ContinueFrt12Record : RowDataRecord
    {
        public const short sid = 0x87F;
        public ContinueFrt12Record(RecordInputStream ris)
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

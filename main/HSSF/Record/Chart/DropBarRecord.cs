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
    /// The DropBar record specifies the attributes of the up bars or the down bars between multiple 
    /// series of a line chart group and specifies the beginning of a collection of records as 
    /// defined by the Chart Sheet Substream ABNF. The first of these collections in the line chart 
    /// group specifies the attributes of the up bars. The second specifies the attributes of the 
    /// down bars. If this record exists, then the chart group type MUST be line and the field cSer 
    /// in the record SeriesList MUST be greater than 1.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class DropBarRecord : RowDataRecord
    {
        public const short sid = 0x103D;
        public DropBarRecord(RecordInputStream ris)
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

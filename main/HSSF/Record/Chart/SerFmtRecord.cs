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
    /// The SerFmt record specifies properties of the associated data points, data markers, 
    /// or lines of the series. The associated data points, data markers, or lines of the 
    /// series are specified by the preceding DataFormat record. If this record is not 
    /// present in the sequence of records that conforms to the SS rule of the Chart Sheet 
    /// Substream ABNF, then the properties of the associated data points, data markers, 
    /// or lines of the series are specified by the default values of the fields of this record.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class SerFmtRecord : RowDataRecord
    {
         public const short sid = 0x105D;
         public SerFmtRecord(RecordInputStream ris)
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

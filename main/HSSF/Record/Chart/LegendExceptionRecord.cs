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
using System;

namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// The LegendException record specifies information about a legend entry which was 
    /// changed from the default legend entry settings, and specifies the beginning of 
    /// a collection of records as defined by the Chart Sheet Substream ABNF. 
    /// The collection of records specifies legend entry formatting. On a chart where 
    /// the legend contains legend entries for the series and trendlines, as defined 
    /// in the legend overview, there MUST be zero instances or one instance of this 
    /// record in the sequence of records that conform to the SERIESFORMAT rule.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class LegendExceptionRecord : RowDataRecord
    {
        public const short sid = 0x1043;
        public LegendExceptionRecord(RecordInputStream ris)
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
        public short LegendEntry
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }
    }
}

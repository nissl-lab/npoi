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
    /// The RichTextStream record specifies additional text properties for the text in 
    /// the entire chart, text in the current legend, text in the current legend entry, 
    /// or text in the attached label. These text properties are a superset of the 
    /// properties stored in the Text, Font, FontX, BRAI, and ObjectLink records based 
    /// on the following table, as specified by the Chart Sheet Substream ABNF. In each 
    /// case, the associated Font record is specified by the associated FontX record. 
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class RichTextStreamRecord : RowDataRecord
    {
        public const short sid = 0x8a6;
        public RichTextStreamRecord(RecordInputStream ris)
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

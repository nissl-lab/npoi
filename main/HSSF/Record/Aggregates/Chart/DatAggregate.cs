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

using NPOI.HSSF.Record.Chart;
using NPOI.HSSF.Model;
using System.Diagnostics;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// DAT = Dat Begin LD End
    /// </summary>
    public class DatAggregate : ChartRecordAggregate
    {
        private DatRecord dat = null;
        private LDAggregate ld = null;

        public DatAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_DAT, container)
        {
            dat = (DatRecord)rs.GetNext();
            rs.GetNext();
            ld = new LDAggregate(rs, this);

            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(dat);
            rv.VisitRecord(BeginRecord.instance);
            ld.VisitContainedRecords(rv);
            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

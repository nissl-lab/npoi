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

using System.Collections.Generic;
using NPOI.HSSF.Record.Chart;
using NPOI.HSSF.Model;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// CRTMLFRT = CrtMlFrt *CrtMlFrtContinue
    /// </summary>
    public class CrtMlFrtAggregate : ChartRecordAggregate
    {
        private CrtMlFrtRecord crtmlFrt = null;
        private List<CrtMlFrtContinueRecord> continues = new List<CrtMlFrtContinueRecord>();
        public CrtMlFrtAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_CRTMLFRT, container)
        {
            crtmlFrt = (CrtMlFrtRecord)rs.GetNext();
            if (rs.PeekNextChartSid() == CrtMlFrtContinueRecord.sid)
            {
                while (rs.PeekNextChartSid() == CrtMlFrtContinueRecord.sid)
                {
                    continues.Add((CrtMlFrtContinueRecord)rs.GetNext());
                }
            }
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            WriteStartBlock(rv);
            rv.VisitRecord(crtmlFrt);
            foreach (CrtMlFrtContinueRecord cr in continues)
                rv.VisitRecord(cr);
        }
    }
}

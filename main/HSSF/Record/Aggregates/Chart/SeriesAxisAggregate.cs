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

using NPOI.HSSF.Model;
using System.Diagnostics;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// SERIESAXIS = Axis Begin [CatSerRange] AXS [CRTMLFRT] End
    /// </summary>
    public class SeriesAxisAggregate: ChartRecordAggregate
    {
        private AxisRecord axis;
        private CatSerRangeRecord catSerRange;
        private AXSAggregate axs;
        private CrtMlFrtAggregate crtmlfrt;
        public SeriesAxisAggregate(RecordStream rs, ChartRecordAggregate container)
            : base("SERIESAXIS", container)
        {
            axis = (AxisRecord)rs.GetNext();
            rs.GetNext();

            if (rs.PeekNextChartSid() == CatSerRangeRecord.sid)
                catSerRange = (CatSerRangeRecord)rs.GetNext();

            axs = new AXSAggregate(rs, this);
            if (rs.PeekNextChartSid() == CrtMlFrtRecord.sid)
                crtmlfrt = new CrtMlFrtAggregate(rs, this);

            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(axis);
            rv.VisitRecord(BeginRecord.instance);

            if (catSerRange != null)
                rv.VisitRecord(catSerRange);

            axs.VisitContainedRecords(rv);
            if (crtmlfrt != null)
                crtmlfrt.VisitContainedRecords(rv);

            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

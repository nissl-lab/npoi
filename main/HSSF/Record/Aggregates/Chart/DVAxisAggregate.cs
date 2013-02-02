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
using NPOI.HSSF.Record.Chart;
using System.Diagnostics;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// DVAXIS = Axis Begin [ValueRange] [AXM] AXS [CRTMLFRT] End
    /// </summary>
    public class DVAxisAggregate: ChartRecordAggregate
    {
        private AxisRecord axis;
        private ValueRangeRecord valueRange;
        private AXMAggregate axm;
        private AXSAggregate axs;
        private CrtMlFrtAggregate crtmlfrt;

        public AxisRecord Axis
        {
            get { return axis; }
        }
        public DVAxisAggregate(RecordStream rs, ChartRecordAggregate container, AxisRecord axis)
            : base(RuleName_DVAXIS, container)
        {
            if (axis == null)
            {
                this.axis = (AxisRecord)rs.GetNext();
                rs.GetNext();
            }
            else
            {
                this.axis = axis;
            }

            if (rs.PeekNextChartSid() == ValueRangeRecord.sid)
                valueRange = (ValueRangeRecord)rs.GetNext();

            if (rs.PeekNextChartSid() == YMultRecord.sid)
                axm = new AXMAggregate(rs, this);

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
            if (valueRange != null)
                rv.VisitRecord(valueRange);
            if (axm != null)
                axm.VisitContainedRecords(rv);

            axs.VisitContainedRecords(rv);

            if (crtmlfrt != null)
                crtmlfrt.VisitContainedRecords(rv);

            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

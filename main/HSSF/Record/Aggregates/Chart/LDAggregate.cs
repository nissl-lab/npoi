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
    /// LD = Legend Begin Pos ATTACHEDLABEL [FRAME] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End
    /// </summary>
    public class LDAggregate : ChartRecordAggregate
    {
        private LegendRecord legend = null;
        private PosRecord pos = null;
        private AttachedLabelAggregate attachedLabel = null;
        private FrameAggregate frame = null;
        private CrtLayout12Record crtLayout = null;
        private TextPropsAggregate textProps = null;
        private CrtMlFrtAggregate crtMlFrt = null;
        public LDAggregate(RecordStream rs, ChartRecordAggregate container)
            :base(RuleName_LD, container)
        {
            legend = (LegendRecord)rs.GetNext();
            rs.GetNext();//BeginRecord
            pos = (PosRecord)rs.GetNext();
            attachedLabel = new AttachedLabelAggregate(rs, this);

            if (rs.PeekNextChartSid() == FrameRecord.sid)
            {
                frame = new FrameAggregate(rs, this);
            }
            if (rs.PeekNextChartSid() == CrtLayout12Record.sid)
            {
                crtLayout = (CrtLayout12Record)rs.GetNext();
            }
            if (rs.PeekNextChartSid() == TextPropsStreamRecord.sid)
            {
                textProps = new TextPropsAggregate(rs, this);
            }
            if (rs.PeekNextChartSid() == CrtMlFrtRecord.sid)
            {
                crtMlFrt = new CrtMlFrtAggregate(rs, this);
            }
            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(legend);
            rv.VisitRecord(BeginRecord.instance);
            rv.VisitRecord(pos);
            attachedLabel.VisitContainedRecords(rv);

            if (frame != null)
                frame.VisitContainedRecords(rv);

            if (crtLayout != null)
                rv.VisitRecord(crtLayout);

            if (textProps != null)
                textProps.VisitContainedRecords(rv);

            if (crtMlFrt != null)
                crtMlFrt.VisitContainedRecords(rv);
            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

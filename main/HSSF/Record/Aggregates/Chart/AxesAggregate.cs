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
using System.Collections.Generic;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;
using System.Diagnostics;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// AXES = [IVAXIS DVAXIS [SERIESAXIS] / DVAXIS DVAXIS] *3ATTACHEDLABEL [PlotArea FRAME]
    /// </summary>
    public class AxesAggregate : ChartRecordAggregate
    {
        private IVAxisAggregate ivaxis;
        private DVAxisAggregate dvaxis;

        private DVAxisAggregate dvaxisSecond;
        private SeriesAxisAggregate seriesAxis;

        private List<AttachedLabelAggregate> attachedLabelList = new List<AttachedLabelAggregate>();
        private PlotAreaRecord plotArea;
        private FrameAggregate frame;

        public AxesAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_AXES, container)
        {
            if (rs.PeekNextChartSid() == AxisRecord.sid)
            {
                AxisRecord axis = (AxisRecord)rs.GetNext();
                rs.GetNext();
                int sid = rs.PeekNextChartSid();
                if (sid == CatSerRangeRecord.sid)
                {
                    ivaxis = new IVAxisAggregate(rs, this, axis);
                }
                else if (sid == ValueRangeRecord.sid)
                {
                    dvaxis = new DVAxisAggregate(rs, this, axis);
                }
                else
                    throw new InvalidOperationException(string.Format("Invalid record sid=0x{0:X}. Shoud be CatSerRangeRecord or ValueRangeRecord", sid));

                Debug.Assert(rs.PeekNextChartSid() == AxisRecord.sid);
                dvaxisSecond = new DVAxisAggregate(rs, this, null);
                if (rs.PeekNextChartSid() == AxisRecord.sid)
                    seriesAxis = new SeriesAxisAggregate(rs, this);

                while (rs.PeekNextChartSid() == TextRecord.sid)
                {
                    attachedLabelList.Add(new AttachedLabelAggregate(rs, this));
                }
                if (rs.PeekNextChartSid() == PlotAreaRecord.sid)
                {
                    plotArea = (PlotAreaRecord)rs.GetNext();
                    if (rs.PeekNextChartSid() == FrameRecord.sid)
                        frame = new FrameAggregate(rs, this);
                }
            }
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            if (ivaxis != null)
                ivaxis.VisitContainedRecords(rv);
            if (dvaxis != null)
                dvaxis.VisitContainedRecords(rv);
            dvaxisSecond.VisitContainedRecords(rv);
            if (seriesAxis != null)
                seriesAxis.VisitContainedRecords(rv);

            foreach (AttachedLabelAggregate al in attachedLabelList)
                al.VisitContainedRecords(rv);
            if (plotArea != null)
            {
                rv.VisitRecord(plotArea);
                if (frame != null)
                    frame.VisitContainedRecords(rv);
            }
        }
    }
}

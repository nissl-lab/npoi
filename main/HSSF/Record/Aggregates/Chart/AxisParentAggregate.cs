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
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;
using System.Diagnostics;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// AXISPARENT = AxisParent Begin Pos [AXES] 1*4CRT End
    /// </summary>
    public class AxisParentAggregate: ChartRecordAggregate
    {
        private AxisParentRecord axisPraent = null;
        private PosRecord pos = null;
        private AxesAggregate axes = null;
        private List<CRTAggregate> crtList = new List<CRTAggregate>();

        public AxisParentRecord AxisParent
        {
            get { return axisPraent; }
        }

        public AxisParentAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_AXISPARENT, container)
        {
            axisPraent = (AxisParentRecord)rs.GetNext();
            rs.GetNext();
            pos = (PosRecord)rs.GetNext();
            if (ChartFormatRecord.sid != rs.PeekNextChartSid())
            {
                try
                {
                    axes = new AxesAggregate(rs, this);
                }
                catch
                {
                    Debug.Print("not find axes rule records");
                    axes = null;
                }
            }
            Debug.Assert(ChartFormatRecord.sid == rs.PeekNextChartSid());
            while (ChartFormatRecord.sid == rs.PeekNextChartSid())
            {
                crtList.Add(new CRTAggregate(rs, this));
            }
            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(axisPraent);
            rv.VisitRecord(BeginRecord.instance);
            rv.VisitRecord(pos);
            if (axes != null)
                axes.VisitContainedRecords(rv);
            foreach (CRTAggregate crt in crtList)
                crt.VisitContainedRecords(rv);

            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

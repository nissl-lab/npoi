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
    /// IVAXIS = Axis Begin [CatSerRange] AxcExt [CatLab] AXS [CRTMLFRT] End
    /// </summary>
    public class IVAxisAggregate: ChartRecordAggregate
    {
        private AxisRecord axis;
        private CatSerRangeRecord catSerRange;
        private AxcExtRecord axcExt;
        private CatLabRecord catLab;
        private AXSAggregate axs;
        //more than one CrtMlFrtRecord?
        private List<CrtMlFrtAggregate> crtmlfrtList = new List<CrtMlFrtAggregate>();


        public IVAxisAggregate(RecordStream rs, ChartRecordAggregate container, AxisRecord axis)
            : base(RuleName_IVAXIS, container)
        {
            this.axis = axis;

            if (rs.PeekNextChartSid() == CatSerRangeRecord.sid)
                catSerRange = (CatSerRangeRecord)rs.GetNext();

            Debug.Assert(rs.PeekNextChartSid() == AxcExtRecord.sid);
            axcExt = (AxcExtRecord)rs.GetNext();

            if (rs.PeekNextChartSid() == CatLabRecord.sid)
                catLab = (CatLabRecord)rs.GetNext();

            axs = new AXSAggregate(rs, this);
            while (rs.PeekNextChartSid() == CrtMlFrtRecord.sid)
                crtmlfrtList.Add(new CrtMlFrtAggregate(rs, this));

            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(axis);
            rv.VisitRecord(BeginRecord.instance);

            if (catSerRange != null)
                rv.VisitRecord(catSerRange);

            rv.VisitRecord(axcExt);
            if (catLab != null)
            {
                WriteStartBlock(rv);
                rv.VisitRecord(catLab);
            }
            axs.VisitContainedRecords(rv);

            foreach (CrtMlFrtAggregate crtmlfrt in crtmlfrtList)
                crtmlfrt.VisitContainedRecords(rv);

            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

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

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// SHAPEPROPS = ShapePropsStream *ContinueFrt12
    /// </summary>
    public class ShapePropsAggregate : ChartRecordAggregate
    {
        private ShapePropsStreamRecord shapProps = null;
        private List<ContinueFrt12Record> continues = new List<ContinueFrt12Record>();
        public ShapePropsAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_SHAPEPROPS, container)
        {
            shapProps = (ShapePropsStreamRecord)rs.GetNext();
            if (rs.PeekNextChartSid() == CrtMlFrtContinueRecord.sid)
            {
                while (rs.PeekNextChartSid() == CrtMlFrtContinueRecord.sid)
                {
                    continues.Add((ContinueFrt12Record)rs.GetNext());
                }
            }
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            WriteStartBlock(rv);
            rv.VisitRecord(shapProps);
            foreach (ContinueFrt12Record cr in continues)
                rv.VisitRecord(cr);
        }
    }
}

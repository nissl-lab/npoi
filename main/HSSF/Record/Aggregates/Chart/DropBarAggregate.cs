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
    /// DROPBAR = DropBar Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End
    /// </summary>
    public class DropBarAggregate : ChartRecordAggregate
    {
        private DropBarRecord dropBar = null;
        private LineFormatRecord lineFormat = null;
        private AreaFormatRecord areaFormat = null;
        private GelFrameAggregate gelFrame = null;
        private ShapePropsAggregate shapProps = null;
        public DropBarAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_DROPBAR, container)
        {
            dropBar = (DropBarRecord)rs.GetNext();
            rs.GetNext();
            lineFormat = (LineFormatRecord)rs.GetNext();
            areaFormat = (AreaFormatRecord)rs.GetNext();

            if (rs.PeekNextChartSid() == GelFrameRecord.sid)
            {
                gelFrame = new GelFrameAggregate(rs, this);
            }
            if (rs.PeekNextChartSid() == ShapePropsStreamRecord.sid)
            {
                shapProps = new ShapePropsAggregate(rs, this);
            }

            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(dropBar);
            rv.VisitRecord(BeginRecord.instance);
            rv.VisitRecord(lineFormat);
            rv.VisitRecord(areaFormat);

            if (gelFrame != null)
                gelFrame.VisitContainedRecords(rv);
            if (shapProps != null)
                shapProps.VisitContainedRecords(rv);

            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

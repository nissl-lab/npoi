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
namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// AXM = YMult StartObject ATTACHEDLABEL EndObject
    /// </summary>
    public class AXMAggregate : ChartRecordAggregate
    {
        private YMultRecord yMult = null;
        private ChartStartObjectRecord startObject = null;
        private AttachedLabelAggregate attachedLabel = null;
        private ChartEndObjectRecord endObject = null;

        public AXMAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_AXM, container)
        {
            yMult = (YMultRecord)rs.GetNext();
            startObject = (ChartStartObjectRecord)rs.GetNext();
            attachedLabel = new AttachedLabelAggregate(rs, this);
            
            endObject = (ChartEndObjectRecord)rs.GetNext();
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            WriteStartBlock(rv);//??
            rv.VisitRecord(yMult);
            rv.VisitRecord(startObject);
            IsInStartObject = true;
            attachedLabel.VisitContainedRecords(rv);
            IsInStartObject = false;
            rv.VisitRecord(endObject);
        }
    }
}

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
    /// DFTTEXT = [DataLabExt StartObject] DefaultText ATTACHEDLABEL [EndObject]
    /// </summary>
    public class DFTTextAggregate : ChartRecordAggregate
    {
        private DataLabExtRecord dataLabExt = null;
        private ChartStartObjectRecord startObject = null;
        private DefaultTextRecord defaultText = null;
        private AttachedLabelAggregate attachedLabel = null;
        private ChartEndObjectRecord endObject = null;

        public DefaultTextRecord DefaultText
        {
            get { return defaultText; }
        }

        public DFTTextAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_DFTTEXT, container)
        {
            if (rs.PeekNextChartSid() == DataLabExtRecord.sid)
            {
                dataLabExt = (DataLabExtRecord)rs.GetNext();
                startObject = (ChartStartObjectRecord)rs.GetNext();
            }
            defaultText = (DefaultTextRecord)rs.GetNext();
            attachedLabel = new AttachedLabelAggregate(rs, this);

            if (startObject != null)
                endObject = (ChartEndObjectRecord)rs.GetNext();
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            if (dataLabExt != null)
            {
                rv.VisitRecord(dataLabExt);
                rv.VisitRecord(startObject);
                IsInStartObject = true;
            }

            rv.VisitRecord(defaultText);
            attachedLabel.VisitContainedRecords(rv);

            if (endObject != null)
            {
                rv.VisitRecord(endObject);
                IsInStartObject = false;
            }
        }
    }
}

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
    /// SERIESDATA = Dimensions 3(SIIndex *(Number / BoolErr / Blank / Label))
    /// </summary>
    public class SeriesDataAggregate : RecordAggregate
    {
        private DimensionsRecord dimensions = null;
        Dictionary<SeriesIndexRecord, List<Record>> dicData = new Dictionary<SeriesIndexRecord, List<Record>>();
        public SeriesDataAggregate(RecordStream rs)
        {
            dimensions = (DimensionsRecord)rs.GetNext();
            while (rs.PeekNextChartSid() == SeriesIndexRecord.sid)
            {
                SeriesIndexRecord siIndex = (SeriesIndexRecord)rs.GetNext();
                int sid = rs.PeekNextChartSid();
                List<Record> dataList = new List<Record>();
                while (sid == NumberRecord.sid || sid == BoolErrRecord.sid || sid == BlankRecord.sid || sid == LabelRecord.sid)
                {
                    dataList.Add(rs.GetNext());
                }
                dicData.Add(siIndex, dataList);
            }
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(dimensions);
            foreach (KeyValuePair<SeriesIndexRecord, List<Record>> kv in dicData)
            {
                rv.VisitRecord(kv.Key);
                foreach (Record r in kv.Value)
                    rv.VisitRecord(r);
            }
        }
    }
}

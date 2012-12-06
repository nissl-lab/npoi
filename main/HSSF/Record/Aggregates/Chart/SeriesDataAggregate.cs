using System;
using System.Collections.Generic;
using System.Text;
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
            while (rs.PeekNextSid() == SeriesIndexRecord.sid)
            {
                SeriesIndexRecord siIndex = (SeriesIndexRecord)rs.GetNext();
                int sid = rs.PeekNextSid();
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

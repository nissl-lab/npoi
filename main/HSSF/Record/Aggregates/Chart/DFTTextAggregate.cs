using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;
namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// DFTTEXT = [DataLabExt StartObject] DefaultText ATTACHEDLABEL [EndObject]
    /// </summary>
    public class DFTTextAggregate : RecordAggregate
    {
        private DataLabExtRecord dataLabExt = null;
        private ChartStartObjectRecord startObject = null;
        private DefaultTextRecord defaultText = null;
        private AttachedLabelAggregate attachedLabel = null;
        private ChartEndObjectRecord endObject = null;

        public DFTTextAggregate(RecordStream rs)
        {
            if (rs.PeekNextSid() == DataLabExtRecord.sid)
            {
                dataLabExt = (DataLabExtRecord)rs.GetNext();
                startObject = (ChartStartObjectRecord)rs.GetNext();
            }
            defaultText = (DefaultTextRecord)rs.GetNext();
            attachedLabel = new AttachedLabelAggregate(rs);

            if (startObject != null)
                endObject = (ChartEndObjectRecord)rs.GetNext();
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            if (dataLabExt != null)
            {
                rv.VisitRecord(dataLabExt);
                rv.VisitRecord(startObject);
            }

            rv.VisitRecord(defaultText);
            attachedLabel.VisitContainedRecords(rv);

            if (endObject != null)
                rv.VisitRecord(endObject);
        }
    }
}

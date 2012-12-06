using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;
namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// AXM = YMult StartObject ATTACHEDLABEL EndObject
    /// </summary>
    public class AXMAggregate : RecordAggregate
    {
        private YMultRecord yMult = null;
        private ChartStartObjectRecord startObject = null;
        private AttachedLabelAggregate attachedLabel = null;
        private ChartEndObjectRecord endObject = null;

        public AXMAggregate(RecordStream rs)
        {
            yMult = (YMultRecord)rs.GetNext();
            startObject = (ChartStartObjectRecord)rs.GetNext();

            attachedLabel = new AttachedLabelAggregate(rs);

            endObject = (ChartEndObjectRecord)rs.GetNext();
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(yMult);
            rv.VisitRecord(startObject);
            attachedLabel.VisitContainedRecords(rv);
            rv.VisitRecord(endObject);
        }
    }
}

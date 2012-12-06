using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// SHAPEPROPS = ShapePropsStream *ContinueFrt12
    /// </summary>
    public class ShapePropsAggregate : RecordAggregate
    {
        private ShapePropsStreamRecord crtmlFrt = null;
        private List<ContinueFrt12Record> continues = new List<ContinueFrt12Record>();
        public ShapePropsAggregate(RecordStream rs)
        {
            crtmlFrt = (ShapePropsStreamRecord)rs.GetNext();
            if (rs.PeekNextSid() == CrtMlFrtContinueRecord.sid)
            {
                while (rs.PeekNextSid() == CrtMlFrtContinueRecord.sid)
                {
                    continues.Add((ContinueFrt12Record)rs.GetNext());
                }
            }
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(crtmlFrt);
            foreach (ContinueFrt12Record cr in continues)
                rv.VisitRecord(cr);
        }
    }
}

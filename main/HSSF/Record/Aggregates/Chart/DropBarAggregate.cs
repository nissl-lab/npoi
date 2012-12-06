using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;
namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// DROPBAR = DropBar Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End
    /// </summary>
    public class DropBarAggregate : RecordAggregate
    {
        private DropBarRecord dropBar = null;
        private LineFormatRecord lineFormat = null;
        private AreaFormatRecord areaFormat = null;
        private GelFrameAggregate gelFrame = null;
        private ShapePropsAggregate shapProps = null;
        public DropBarAggregate(RecordStream rs)
        {
            dropBar = (DropBarRecord)rs.GetNext();
            rs.GetNext();
            lineFormat = (LineFormatRecord)rs.GetNext();
            areaFormat = (AreaFormatRecord)rs.GetNext();

            if (rs.PeekNextSid() == GelFrameRecord.sid)
            {
                gelFrame = new GelFrameAggregate(rs);
            }
            if (rs.PeekNextSid() == ShapePropsStreamRecord.sid)
            {
                shapProps = new ShapePropsAggregate(rs);
            }

            rs.GetNext();
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

            rv.VisitRecord(EndRecord.instance);
        }
    }
}

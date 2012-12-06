using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Record.Chart;
using NPOI.HSSF.Model;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// Frame Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End
    /// </summary>
    public class FrameAggregate: RecordAggregate
    {
        private FrameRecord frame = null;
        private LineFormatRecord lineFormat = null;
        private AreaFormatRecord areaFormat = null;
        private GelFrameAggregate gelFrame = null;
        private ShapePropsAggregate shapeProps = null;
        public FrameAggregate(RecordStream rs)
        {
            frame = (FrameRecord)rs.GetNext();
            rs.GetNext();//BeginRecord
            lineFormat = (LineFormatRecord)rs.GetNext();
            areaFormat = (AreaFormatRecord)rs.GetNext();
            if (rs.PeekNextSid() == GelFrameRecord.sid)
            {
                gelFrame = new GelFrameAggregate(rs);
            }
            if (rs.PeekNextSid() == ShapePropsStreamRecord.sid)
            {
                shapeProps = new ShapePropsAggregate(rs);
            }
            rs.GetNext();//EndRecord
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(frame);
            rv.VisitRecord(BeginRecord.instance);
            rv.VisitRecord(lineFormat);
            rv.VisitRecord(areaFormat);
            if (gelFrame != null)
                gelFrame.VisitContainedRecords(rv);
            if (shapeProps != null)
                shapeProps.VisitContainedRecords(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

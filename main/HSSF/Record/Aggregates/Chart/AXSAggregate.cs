using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// AXS = [IFmtRecord] [Tick] [FontX] *4(AxisLine LineFormat) [AreaFormat] 
    /// [GELFRAME] *4SHAPEPROPS [TextPropsStream *ContinueFrt12]
    /// </summary>
    public class AXSAggregate : RecordAggregate
    {
        private IFmtRecordRecord ifmt = null;
        private TickRecord tick = null;
        private FontXRecord fontx = null;
        private List<AxisLineRecord> axisLines = new List<AxisLineRecord>();
        private List<LineFormatRecord> lineFormats = new List<LineFormatRecord>();
        private AreaFormatRecord areaFormat = null;
        private GelFrameAggregate gelFrame = null;
        private List<ShapePropsAggregate> shapes = new List<ShapePropsAggregate>();
        private TextPropsStreamRecord textProps = null;
        private List<ContinueFrt12Record> continues = new List<ContinueFrt12Record>();

        public AXSAggregate(RecordStream rs)
        {
            if (rs.PeekNextSid() == IFmtRecordRecord.sid)
                ifmt = (IFmtRecordRecord)rs.GetNext();
            if (rs.PeekNextSid() == IFmtRecordRecord.sid)
                tick = (TickRecord)rs.GetNext();
            if (rs.PeekNextSid() == FontXRecord.sid)
                fontx = (FontXRecord)rs.GetNext();
            if (rs.PeekNextSid() == AxisLineRecord.sid)
            {
                while (rs.PeekNextSid() == AxisLineRecord.sid)
                {
                    axisLines.Add((AxisLineRecord)rs.GetNext());
                    lineFormats.Add((LineFormatRecord)rs.GetNext());
                }
            }

            if (rs.PeekNextSid() == AreaFormatRecord.sid)
                areaFormat = (AreaFormatRecord)rs.GetNext();

            if (rs.PeekNextSid() == GelFrameRecord.sid)
                gelFrame = new GelFrameAggregate(rs);
            if (rs.PeekNextSid() == ShapePropsStreamRecord.sid)
            {
                while (rs.PeekNextSid() == ShapePropsStreamRecord.sid)
                {
                    shapes.Add(new ShapePropsAggregate(rs));
                }
            }
            if (rs.PeekNextSid() == TextPropsStreamRecord.sid)
            {
                textProps = (TextPropsStreamRecord)rs.GetNext();
                while (rs.PeekNextSid() == ContinueFrt12Record.sid)
                {
                    continues.Add((ContinueFrt12Record)rs.GetNext());
                }
            }
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            if (ifmt != null) rv.VisitRecord(ifmt);
            if (tick != null) rv.VisitRecord(tick);
            if (fontx != null) rv.VisitRecord(fontx);
            for (int i = 0; i < axisLines.Count; i++)
            {
                rv.VisitRecord(axisLines[i]);
                rv.VisitRecord(lineFormats[i]);
            }

            if (areaFormat != null) rv.VisitRecord(areaFormat);
            if (gelFrame != null) gelFrame.VisitContainedRecords(rv);
            foreach (ShapePropsAggregate shape in shapes)
                shape.VisitContainedRecords(rv);
            if (textProps != null)
            {
                rv.VisitRecord(textProps);
                foreach (ContinueFrt12Record c in continues)
                    rv.VisitRecord(c);
            }
        }
    }
}

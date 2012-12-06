using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// SS = DataFormat Begin [Chart3DBarShape] [LineFormat AreaFormat PieFormat] [SerFmt] 
    /// [GELFRAME] [MarkerFormat] [AttachedLabel] *2SHAPEPROPS [CRTMLFRT] End
    /// </summary>
    public class SSAggregate : RecordAggregate
    {
        private DataFormatRecord dataFormat = null;
        private Chart3DBarShapeRecord chart3DBarShape = null;
        private LineFormatRecord lineFormat = null;
        private AreaFormatRecord areaFormat = null;
        private PieFormatRecord pieFormat = null;
        private SerFmtRecord serFmt = null;
        private GelFrameAggregate gelFrame = null;
        private MarkerFormatRecord markerFormat = null;
        private AttachedLabelRecord attachedLabel = null;
        private ShapePropsAggregate shapeProps1 = null;
        private ShapePropsAggregate shapeProps2 = null;
        private CrtMlFrtAggregate crtMlFrt = null;
        public SSAggregate(RecordStream rs)
        {
            dataFormat = (DataFormatRecord)rs.GetNext();
            rs.GetNext();
            if (rs.PeekNextSid() == Chart3DBarShapeRecord.sid)
                chart3DBarShape = (Chart3DBarShapeRecord)rs.GetNext();
            if (rs.PeekNextSid() == LineFormatRecord.sid)
            {
                lineFormat = (LineFormatRecord)rs.GetNext();
                areaFormat = (AreaFormatRecord)rs.GetNext();
                pieFormat = (PieFormatRecord)rs.GetNext();
            }
            if(rs.PeekNextSid()==SerFmtRecord.sid)
                serFmt = (SerFmtRecord)rs.GetNext();

            if (rs.PeekNextSid() == GelFrameRecord.sid)
                gelFrame = new GelFrameAggregate(rs);

            if (rs.PeekNextSid() == MarkerFormatRecord.sid)
                markerFormat = (MarkerFormatRecord)rs.GetNext();

            if (rs.PeekNextSid() == AttachedLabelRecord.sid)
                attachedLabel = (AttachedLabelRecord)rs.GetNext();

            if (rs.PeekNextSid() == ShapePropsStreamRecord.sid)
                shapeProps1 = new ShapePropsAggregate(rs);

            if (rs.PeekNextSid() == ShapePropsStreamRecord.sid)
                shapeProps2 = new ShapePropsAggregate(rs);

            if (rs.PeekNextSid() == CrtMlFrtRecord.sid)
                crtMlFrt = new CrtMlFrtAggregate(rs);

            rs.GetNext();
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(dataFormat);
            rv.VisitRecord(BeginRecord.instance);

            if (chart3DBarShape != null)
                rv.VisitRecord(chart3DBarShape);
            if (lineFormat != null)
            {
                rv.VisitRecord(lineFormat);
                rv.VisitRecord(areaFormat);
                rv.VisitRecord(pieFormat);
            }

            if (serFmt != null)
                rv.VisitRecord(serFmt);

            if (gelFrame != null)
                gelFrame.VisitContainedRecords(rv);
            if (markerFormat != null)
                rv.VisitRecord(markerFormat);

            if (attachedLabel != null)
                rv.VisitRecord(attachedLabel);
            if (shapeProps1 != null)
                shapeProps1.VisitContainedRecords(rv);
            if (shapeProps2 != null)
                shapeProps2.VisitContainedRecords(rv);
            if (crtMlFrt != null)
                crtMlFrt.VisitContainedRecords(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

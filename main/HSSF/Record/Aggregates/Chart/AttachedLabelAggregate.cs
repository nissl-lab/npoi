using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// ATTACHEDLABEL = Text Begin Pos [FontX] [AlRuns] AI [FRAME] [ObjectLink] [DataLabExtContents] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End
    /// AI = BRAI [SeriesText]
    /// </summary>
    public class AttachedLabelAggregate : RecordAggregate
    {
        private TextRecord text = null;
        private PosRecord pos = null;
        private FontXRecord fontX = null;
        private AlRunsRecord alRuns = null;
        private BRAIRecord brai = null;
        private SeriesTextRecord seriesText = null;
        private FrameAggregate frame = null;
        private ObjectLinkRecord objectLink = null;
        private DataLabExtContentsRecord dataLab = null;
        private CrtLayout12Record crtLayout = null;
        private TextPropsAggregate textProps = null;
        private CrtMlFrtAggregate crtMlFrt = null;
        public AttachedLabelAggregate(RecordStream rs)
        {
            text = (TextRecord)rs.GetNext();
            rs.GetNext();//BeginRecord
            pos = (PosRecord)rs.GetNext();
            if (rs.PeekNextSid() == FontXRecord.sid)
            {
                fontX = (FontXRecord)rs.GetNext();
            }
            if (rs.PeekNextSid() == AlRunsRecord.sid)
            {
                alRuns = (AlRunsRecord)rs.GetNext();
            }
            brai = (BRAIRecord)rs.GetNext();
            if (rs.PeekNextSid() == SeriesTextRecord.sid)
            {
                seriesText = (SeriesTextRecord)rs.GetNext();
            }
            if (rs.PeekNextSid() == FrameRecord.sid)
            {
                frame = new FrameAggregate(rs);
            }
            if (rs.PeekNextSid() == ObjectLinkRecord.sid)
            {
                objectLink = (ObjectLinkRecord)rs.GetNext();
            }
            if (rs.PeekNextSid() == DataLabExtContentsRecord.sid)
            {
                dataLab = (DataLabExtContentsRecord)rs.GetNext();
            }
            if (rs.PeekNextSid() == CrtLayout12Record.sid)
            {
                crtLayout = (CrtLayout12Record)rs.GetNext();
            }
            if (rs.PeekNextSid() == FrameRecord.sid)
            {
                textProps = new TextPropsAggregate(rs);
            }
            if (rs.PeekNextSid() == FrameRecord.sid)
            {
                crtMlFrt = new CrtMlFrtAggregate(rs);
            }
            rs.GetNext();//EndRecord
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(text);
            rv.VisitRecord(BeginRecord.instance);
            rv.VisitRecord(pos);

            if (fontX != null)
                rv.VisitRecord(fontX);

            if (alRuns != null)
                rv.VisitRecord(alRuns);

            rv.VisitRecord(brai);

            if (seriesText != null)
                rv.VisitRecord(seriesText);

            if (frame != null)
                frame.VisitContainedRecords(rv);

            if (objectLink != null)
                rv.VisitRecord(objectLink);

            if (dataLab != null)
                rv.VisitRecord(dataLab);

            if (crtLayout != null)
                rv.VisitRecord(crtLayout);

            if (textProps != null)
                textProps.VisitContainedRecords(rv);

            if (crtMlFrt != null)
                crtMlFrt.VisitContainedRecords(rv);

            rv.VisitRecord(EndRecord.instance);
        }
    }
}

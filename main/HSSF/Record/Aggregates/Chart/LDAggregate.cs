using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// LD = Legend Begin Pos ATTACHEDLABEL [FRAME] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End
    /// </summary>
    public class LDAggregate : RecordAggregate
    {
        private LegendRecord legend = null;
        private PosRecord pos = null;
        private AttachedLabelAggregate attachedLabel = null;
        private FrameAggregate frame = null;
        private CrtLayout12Record crtLayout = null;
        private TextPropsAggregate textProps = null;
        private CrtMlFrtAggregate crtMlFrt = null;
        public LDAggregate(RecordStream rs)
        {
            legend = (LegendRecord)rs.GetNext();
            rs.GetNext();//BeginRecord
            pos = (PosRecord)rs.GetNext();
            attachedLabel = new AttachedLabelAggregate(rs);

            if (rs.PeekNextSid() == FrameRecord.sid)
            {
                frame = new FrameAggregate(rs);
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
            rs.GetNext(); //EndRecord
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(legend);
            rv.VisitRecord(BeginRecord.instance);
            rv.VisitRecord(pos);
            attachedLabel.VisitContainedRecords(rv);

            if (frame != null)
                frame.VisitContainedRecords(rv);

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

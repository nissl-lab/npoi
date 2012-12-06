using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// SERIESFORMAT = Series Begin 4AI *SS (SerToCrt / (SerParent (SerAuxTrend / SerAuxErrBar)))
    /// *(LegendException [Begin ATTACHEDLABEL [TEXTPROPS] End]) End
    /// </summary>
    public class SeriesFormatAggregate : RecordAggregate
    {
        private SeriesRecord series = null;
        private Dictionary<BRAIRecord, SeriesTextRecord> dic4AI = new Dictionary<BRAIRecord, SeriesTextRecord>();
        private List<SSAggregate> ssList = new List<SSAggregate>();
        private SerToCrtRecord serToCrt = null;
        private SerParentRecord serParent = null;
        private SerAuxTrendRecord serAuxTrend = null;
        private SerAuxErrBarRecord serAuxErrBar = null;
        private List<LegendExceptionAggregate> leList = new List<LegendExceptionAggregate>();
        public SeriesFormatAggregate(RecordStream rs)
        {
            series = (SeriesRecord)rs.GetNext();
            rs.GetNext();
            BRAIRecord ai;
            SeriesTextRecord sText;
            while (rs.PeekNextSid() == BRAIRecord.sid)
            {
                sText = null;
                ai = (BRAIRecord)rs.GetNext();
                if (rs.PeekNextSid() == SeriesTextRecord.sid)
                    sText = (SeriesTextRecord)rs.GetNext();
                dic4AI.Add(ai, sText);
            }
            if (rs.PeekNextSid() == DataFormatRecord.sid)
            {
                while (rs.PeekNextSid() == DataFormatRecord.sid)
                    ssList.Add(new SSAggregate(rs));
            }
            if (rs.PeekNextSid() == SerToCrtRecord.sid)
            {
                serToCrt = (SerToCrtRecord)rs.GetNext();
            }
            else
            {
                if (rs.PeekNextSid() == SerParentRecord.sid)
                {
                    serParent = (SerParentRecord)rs.GetNext();
                    if (rs.PeekNextSid() == SerAuxTrendRecord.sid)
                    {
                        serAuxTrend = (SerAuxTrendRecord)rs.GetNext();
                    }
                    else
                    {
                        serAuxErrBar = (SerAuxErrBarRecord)rs.GetNext();
                    }
                }
            }

            if (rs.PeekNextSid() == LegendExceptionRecord.sid)
            {
                while (rs.PeekNextSid() == LegendExceptionRecord.sid)
                {
                    leList.Add(new LegendExceptionAggregate(rs));
                }
            }

            rs.GetNext();
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(series);
            rv.VisitRecord(BeginRecord.instance);

            foreach(KeyValuePair<BRAIRecord,SeriesTextRecord> kv in dic4AI)
            {
                rv.VisitRecord(kv.Key);
                if (kv.Value != null)
                    rv.VisitRecord(kv.Value);
            }
            foreach (SSAggregate ss in ssList)
                ss.VisitContainedRecords(rv);
            if (serToCrt != null)
            {
                rv.VisitRecord(serToCrt);
            }
            else
            {
                if (serParent != null)
                    rv.VisitRecord(serParent);
                if (serAuxTrend != null)
                    rv.VisitRecord(serAuxTrend);
                else
                    rv.VisitRecord(serAuxErrBar);
            }
            foreach (LegendExceptionAggregate le in leList)
            {
                le.VisitContainedRecords(rv);
            }

            rv.VisitRecord(EndRecord.instance);
        }
        /// <summary>
        /// LegendException [Begin ATTACHEDLABEL [TEXTPROPS] End]
        /// </summary>
        private class LegendExceptionAggregate : RecordAggregate
        {
            private LegendExceptionRecord legendException = null;
            private AttachedLabelAggregate attachedLabel = null;
            private TextPropsAggregate textProps = null;
            public LegendExceptionAggregate(RecordStream rs)
            {
                legendException = (LegendExceptionRecord)rs.GetNext();
                if (rs.PeekNextSid() == BeginRecord.sid)
                {
                    rs.GetNext();
                    attachedLabel = new AttachedLabelAggregate(rs);
                    if (rs.PeekNextSid() == TextPropsStreamRecord.sid || 
                        rs.PeekNextSid() == RichTextStreamRecord.sid)
                    {
                        textProps = new TextPropsAggregate(rs);
                    }
                    rs.GetNext();
                }
            }
            public override void VisitContainedRecords(RecordVisitor rv)
            {
                rv.VisitRecord(legendException);
                if (attachedLabel != null)
                {
                    rv.VisitRecord(BeginRecord.instance);
                    attachedLabel.VisitContainedRecords(rv);
                    if (textProps != null)
                        textProps.VisitContainedRecords(rv);
                    rv.VisitRecord(EndRecord.instance);
                }
            }
        }
    }
}

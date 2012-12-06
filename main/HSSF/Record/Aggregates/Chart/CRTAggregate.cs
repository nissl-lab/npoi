using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// CRT = ChartFormat Begin (Bar / Line / (BopPop [BopPopCustom]) / Pie / Area / Scatter / Radar / 
    /// RadarArea / Surf) CrtLink [SeriesList] [Chart3d] [LD] [2DROPBAR] *4(CrtLine LineFormat) 
    /// *2DFTTEXT [DataLabExtContents] [SS] *4SHAPEPROPS End
    /// </summary>
    public class CRTAggregate : RecordAggregate
    {
        private ChartFormatRecord chartForamt = null;
        private Record chartTypeRecord = null;
        private BopPopCustomRecord bopPopCustom = null;
        private CrtLinkRecord crtLink = null;
        private SeriesListRecord seriesList = null;
        private Chart3dRecord chart3d = null;
        private LDAggregate ld = null;
        private DropBarAggregate dropBar1 = null;
        private DropBarAggregate dropBar2 = null;
        private Dictionary<CrtLineRecord, LineFormatRecord> dicLines = new Dictionary<CrtLineRecord, LineFormatRecord>();
        private DFTTextAggregate dft1 = null;
        private DFTTextAggregate dft2 = null;
        private DataLabExtContentsRecord dataLabExtContents = null;
        private SSAggregate ss = null;
        private List<ShapePropsAggregate> shapeList = new List<ShapePropsAggregate>();

        public CRTAggregate(RecordStream rs)
        {
            chartForamt = (ChartFormatRecord)rs.GetNext();
            rs.GetNext();

            chartTypeRecord = rs.GetNext();
            if (rs.PeekNextSid() == BopPopCustomRecord.sid)
                bopPopCustom = (BopPopCustomRecord)rs.GetNext();
            crtLink = (CrtLinkRecord)rs.GetNext();
            if (rs.PeekNextSid() == SeriesListRecord.sid)
                seriesList = (SeriesListRecord)rs.GetNext();
            if (rs.PeekNextSid() == Chart3dRecord.sid)
                chart3d = (Chart3dRecord)rs.GetNext();
            if (rs.PeekNextSid() == LegendRecord.sid)
                ld = new LDAggregate(rs);
            if (rs.PeekNextSid() == DropBarRecord.sid)
            {
                dropBar1 = new DropBarAggregate(rs);
                dropBar2 = new DropBarAggregate(rs);
            }

            while (rs.PeekNextSid() == CrtLineRecord.sid)
            {
                dicLines.Add((CrtLineRecord)rs.GetNext(), (LineFormatRecord)rs.GetNext());
            }
            if (rs.PeekNextSid() == DataLabExtRecord.sid || rs.PeekNextSid() == DefaultTextRecord.sid)
            {
                dft1 = new DFTTextAggregate(rs);
                if (rs.PeekNextSid() == DataLabExtRecord.sid || rs.PeekNextSid() == DefaultTextRecord.sid)
                {
                    dft2 = new DFTTextAggregate(rs);
                }
            }
            if (rs.PeekNextSid() == DataLabExtContentsRecord.sid)
                dataLabExtContents = (DataLabExtContentsRecord)rs.GetNext();

            if (rs.PeekNextSid() == DataFormatRecord.sid)
                ss = new SSAggregate(rs);
            while (rs.PeekNextSid() == ShapePropsStreamRecord.sid)
                shapeList.Add(new ShapePropsAggregate(rs));

            rs.GetNext();
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(chartForamt);
            rv.VisitRecord(BeginRecord.instance);

            rv.VisitRecord(chartTypeRecord);
            if (bopPopCustom != null)
                rv.VisitRecord(bopPopCustom);
            rv.VisitRecord(crtLink);
            if (seriesList != null)
                rv.VisitRecord(crtLink);
            if (chart3d != null)
                rv.VisitRecord(chart3d);
            if (ld != null)
                ld.VisitContainedRecords(rv);
            if (dropBar1 != null)
            {
                dropBar1.VisitContainedRecords(rv);
                dropBar2.VisitContainedRecords(rv);
            }
            foreach (KeyValuePair<CrtLineRecord, LineFormatRecord> kv in dicLines)
            {
                rv.VisitRecord(kv.Key);
                rv.VisitRecord(kv.Value);
            }
            if (dft1 != null)
            {
                dft1.VisitContainedRecords(rv);
                if (dft2 != null)
                    dft2.VisitContainedRecords(rv);
            }

            if (dataLabExtContents != null)
                rv.VisitRecord(dataLabExtContents);
            if (ss != null)
                ss.VisitContainedRecords(rv);
            foreach (ShapePropsAggregate shape in shapeList)
                shape.VisitContainedRecords(rv);
                

            rv.VisitRecord(EndRecord.instance);
        }
    }
}

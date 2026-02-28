/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

using System.Collections.Generic;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// CRT = ChartFormat Begin (Bar / Line / (BopPop [BopPopCustom]) / Pie / Area / Scatter / Radar / 
    /// RadarArea / Surf) CrtLink [SeriesList] [Chart3d] [LD] [2DROPBAR] *4(CrtLine LineFormat) 
    /// *2DFTTEXT [DataLabExtContents] [SS] *4SHAPEPROPS End
    /// </summary>
    public class CRTAggregate : ChartRecordAggregate
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

        public CRTAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_CRT, container)
        {
            
            chartForamt = (ChartFormatRecord)rs.GetNext();
            rs.GetNext();

            chartTypeRecord = rs.GetNext();
            if (rs.PeekNextChartSid() == BopPopCustomRecord.sid)
                bopPopCustom = (BopPopCustomRecord)rs.GetNext();
            crtLink = (CrtLinkRecord)rs.GetNext();
            if (rs.PeekNextChartSid() == SeriesListRecord.sid)
                seriesList = (SeriesListRecord)rs.GetNext();
            if (rs.PeekNextChartSid() == Chart3dRecord.sid)
                chart3d = (Chart3dRecord)rs.GetNext();
            if (rs.PeekNextChartSid() == LegendRecord.sid)
                ld = new LDAggregate(rs, this);
            if (rs.PeekNextChartSid() == DropBarRecord.sid)
            {
                dropBar1 = new DropBarAggregate(rs, this);
                dropBar2 = new DropBarAggregate(rs, this);
            }

            while (rs.PeekNextChartSid() == CrtLineRecord.sid)
            {
                dicLines.Add((CrtLineRecord)rs.GetNext(), (LineFormatRecord)rs.GetNext());
            }
            if (rs.PeekNextChartSid() == DataLabExtRecord.sid || rs.PeekNextChartSid() == DefaultTextRecord.sid)
            {
                dft1 = new DFTTextAggregate(rs, this);
                if (rs.PeekNextChartSid() == DataLabExtRecord.sid || rs.PeekNextChartSid() == DefaultTextRecord.sid)
                {
                    dft2 = new DFTTextAggregate(rs, this);
                }
            }
            if (rs.PeekNextChartSid() == DataLabExtContentsRecord.sid)
                dataLabExtContents = (DataLabExtContentsRecord)rs.GetNext();

            if (rs.PeekNextChartSid() == DataFormatRecord.sid)
                ss = new SSAggregate(rs, this);
            while (rs.PeekNextChartSid() == ShapePropsStreamRecord.sid)
                shapeList.Add(new ShapePropsAggregate(rs, this));

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

            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

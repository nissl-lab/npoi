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
using System.Diagnostics;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// CHARTFOMATS = Chart Begin *2FONTLIST Scl PlotGrowth [FRAME] *SERIESFORMAT *SS ShtProps 
    /// *2DFTTEXT AxesUsed 1*2AXISPARENT [CrtLayout12A] [DAT] *ATTACHEDLABEL [CRTMLFRT] 
    /// *([DataLabExt StartObject] ATTACHEDLABEL [EndObject]) [TEXTPROPS] *2CRTMLFRT End
    /// </summary>
    public class ChartFormatsAggregate : ChartRecordAggregate
    {
        private ChartRecord chart = null;
        private List<FontListAggregate> fontList = new List<FontListAggregate>();

        private SCLRecord scl = null;
        private PlotGrowthRecord plotGrowth;

        private FrameAggregate frame;
        private List<SeriesFormatAggregate> seriesFormatList = new List<SeriesFormatAggregate>();
        private List<SSAggregate> ssList = new List<SSAggregate>();

        private ShtPropsRecord shtProps;
        private List<DFTTextAggregate> dftTextList = new List<DFTTextAggregate>();
        private AxesUsedRecord axesUsed;

        private AxisParentAggregate axisParent1, axisParent2;
        private CrtLayout12ARecord crt12A;
        private DatAggregate dat;
        private List<AttachedLabelAggregate> attachedLabelList = new List<AttachedLabelAggregate>();
        private CrtMlFrtAggregate crtMlFrt;
        private List<ChartFormatsAttachedLabelAggregate> cfalList = new List<ChartFormatsAttachedLabelAggregate>();
        private TextPropsAggregate textProps;
        private List<CrtMlFrtAggregate> mlfrtList = new List<CrtMlFrtAggregate>();

        public ChartFormatsAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_CHARTFOMATS, container)
        {
            Debug.Assert(rs.PeekNextChartSid() == ChartRecord.sid);
            chart = (ChartRecord)rs.GetNext();
            rs.GetNext();
            while (rs.PeekNextChartSid() == FrtFontListRecord.sid)
            {
                fontList.Add(new FontListAggregate(rs, this));
            }
            Debug.Assert(rs.PeekNextChartSid() == SCLRecord.sid);
            scl = (SCLRecord)rs.GetNext();
            plotGrowth = (PlotGrowthRecord)rs.GetNext();

            if (rs.PeekNextChartSid() == FrameRecord.sid)
            {
                frame = new FrameAggregate(rs, this);
            }

            while (rs.PeekNextChartSid() == SeriesRecord.sid)
            {
                seriesFormatList.Add(new SeriesFormatAggregate(rs, this));
            }

            while (rs.PeekNextChartSid() == DataFormatRecord.sid)
            {
                ssList.Add(new SSAggregate(rs, this));
            }

            Debug.Assert(rs.PeekNextChartSid() == ShtPropsRecord.sid);
            shtProps = (ShtPropsRecord)rs.GetNext();

            while (rs.PeekNextChartSid() == DefaultTextRecord.sid||
                rs.PeekNextChartSid() == DataLabExtRecord.sid)
            {
                dftTextList.Add(new DFTTextAggregate(rs, this));
            }

            Debug.Assert(rs.PeekNextChartSid() == AxesUsedRecord.sid);
            axesUsed = (AxesUsedRecord)rs.GetNext();

            Debug.Assert(rs.PeekNextChartSid() == AxisParentRecord.sid);
            axisParent1 = new AxisParentAggregate(rs, this);
            if (rs.PeekNextChartSid() == AxisParentRecord.sid)
            {
                axisParent2 = new AxisParentAggregate(rs, this);
            }

            if (rs.PeekNextChartSid() == CrtLayout12ARecord.sid)
                crt12A = (CrtLayout12ARecord)rs.GetNext();
            if (rs.PeekNextChartSid() == DatRecord.sid)
                dat = new DatAggregate(rs, container);

            if (rs.PeekNextChartSid() == TextRecord.sid)
            {
                while (rs.PeekNextChartSid() == TextRecord.sid)
                    attachedLabelList.Add(new AttachedLabelAggregate(rs, this));
            }

            if (rs.PeekNextChartSid() == CrtMlFrtRecord.sid)
                crtMlFrt = new CrtMlFrtAggregate(rs, container);

            while (rs.PeekNextChartSid() == DataLabExtRecord.sid || rs.PeekNextChartSid() == TextRecord.sid)
            {
                cfalList.Add(new ChartFormatsAttachedLabelAggregate(rs, this));
            }

            if (rs.PeekNextChartSid() == TextPropsStreamRecord.sid)
                textProps = new TextPropsAggregate(rs, container);

            while (rs.PeekNextChartSid() == CrtMlFrtRecord.sid)
            {
                mlfrtList.Add(new CrtMlFrtAggregate(rs, container));
            }

            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(chart);
            rv.VisitRecord(BeginRecord.instance);
            foreach (FontListAggregate f in fontList)
            {
                f.VisitContainedRecords(rv);
            }
            rv.VisitRecord(scl);
            rv.VisitRecord(plotGrowth);

            if (frame != null)
                frame.VisitContainedRecords(rv);

            short num = 0;
            foreach (SeriesFormatAggregate sf in seriesFormatList)
            {
                sf.VisitContainedRecords(rv);
                sf.SeriesIndex = num++;
            }
            foreach (SSAggregate ss in ssList)
                ss.VisitContainedRecords(rv);
            rv.VisitRecord(shtProps);
            foreach (DFTTextAggregate dft in dftTextList)
                dft.VisitContainedRecords(rv);
            rv.VisitRecord(axesUsed);

            axisParent1.VisitContainedRecords(rv);
            if (axisParent2 != null)
                axisParent2.VisitContainedRecords(rv);

            if (crt12A != null)
                rv.VisitRecord(crt12A);
            if (dat != null)
                dat.VisitContainedRecords(rv);

            foreach (AttachedLabelAggregate al in attachedLabelList)
                al.VisitContainedRecords(rv);

            if (crtMlFrt != null)
                crtMlFrt.VisitContainedRecords(rv);

            foreach (ChartFormatsAttachedLabelAggregate cfa in cfalList)
                cfa.VisitContainedRecords(rv);

            if (textProps != null)
                textProps.VisitContainedRecords(rv);

            foreach (CrtMlFrtAggregate c in mlfrtList)
                c.VisitContainedRecords(rv);

            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }
        private class ChartFormatsAttachedLabelAggregate : ChartRecordAggregate
        {
            private DataLabExtRecord dataLabExt;
            private ChartStartObjectRecord startObject;
            private AttachedLabelAggregate attachedLabel;
            private ChartEndObjectRecord endObject;
            public ChartFormatsAttachedLabelAggregate(RecordStream rs, ChartRecordAggregate container)
                : base("ChartFormatsAttachedLabel", container)
            {
                if (rs.PeekNextChartSid() == DataLabExtRecord.sid)
                {
                    dataLabExt = (DataLabExtRecord)rs.GetNext();
                    startObject = (ChartStartObjectRecord)rs.GetNext();
                }

                attachedLabel = new AttachedLabelAggregate(rs, this);

                if (startObject != null)
                {
                    endObject = (ChartEndObjectRecord)rs.GetNext();
                }
            }

            public override void VisitContainedRecords(RecordVisitor rv)
            {
                if (dataLabExt != null)
                {
                    rv.VisitRecord(dataLabExt);
                    rv.VisitRecord(startObject);
                    IsInStartObject = true;
                }
                attachedLabel.VisitContainedRecords(rv);
                if (endObject != null)
                {
                    IsInStartObject = false;
                    rv.VisitRecord(endObject);
                }
            }
        }
    }
}

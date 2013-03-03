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
    /// SERIESFORMAT = Series Begin 4AI *SS (SerToCrt / (SerParent (SerAuxTrend / SerAuxErrBar)))
    /// *(LegendException [Begin ATTACHEDLABEL [TEXTPROPS] End]) End
    /// </summary>
    public class SeriesFormatAggregate : ChartRecordAggregate
    {
        private SeriesRecord series = null;
        private Dictionary<BRAIRecord, SeriesTextRecord> dic4AI = new Dictionary<BRAIRecord, SeriesTextRecord>();
        private List<SSAggregate> ssList = new List<SSAggregate>();
        private SerToCrtRecord serToCrt = null;
        private SerParentRecord serParent = null;
        private SerAuxTrendRecord serAuxTrend = null;
        private SerAuxErrBarRecord serAuxErrBar = null;
        private List<LegendExceptionAggregate> leList = new List<LegendExceptionAggregate>();
        public SeriesFormatAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_SERIESFORMAT, container)
        {
            series = (SeriesRecord)rs.GetNext();
            rs.GetNext();
            BRAIRecord ai;
            SeriesTextRecord sText;
            while (rs.PeekNextChartSid() == BRAIRecord.sid)
            {
                sText = null;
                ai = (BRAIRecord)rs.GetNext();
                if (rs.PeekNextChartSid() == SeriesTextRecord.sid)
                    sText = (SeriesTextRecord)rs.GetNext();
                dic4AI.Add(ai, sText);
            }
            if (rs.PeekNextChartSid() == DataFormatRecord.sid)
            {
                while (rs.PeekNextChartSid() == DataFormatRecord.sid)
                    ssList.Add(new SSAggregate(rs, this));
            }
            if (rs.PeekNextChartSid() == SerToCrtRecord.sid)
            {
                serToCrt = (SerToCrtRecord)rs.GetNext();
            }
            else
            {
                if (rs.PeekNextChartSid() == SerParentRecord.sid)
                {
                    serParent = (SerParentRecord)rs.GetNext();
                    if (rs.PeekNextChartSid() == SerAuxTrendRecord.sid)
                    {
                        serAuxTrend = (SerAuxTrendRecord)rs.GetNext();
                    }
                    else
                    {
                        serAuxErrBar = (SerAuxErrBarRecord)rs.GetNext();
                    }
                }
            }

            if (rs.PeekNextChartSid() == LegendExceptionRecord.sid)
            {
                while (rs.PeekNextChartSid() == LegendExceptionRecord.sid)
                {
                    leList.Add(new LegendExceptionAggregate(rs, this));
                }
            }

            Record r = rs.GetNext();//EndRecord
            Debug.Assert(r.GetType() == typeof(EndRecord));
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
            WriteEndBlock(rv);
            rv.VisitRecord(EndRecord.instance);
        }

        short _seriesIndex;
        public short SeriesIndex
        {
            get { return _seriesIndex; }
            set { _seriesIndex = value; }
        }
        /// <summary>
        /// LegendException [Begin ATTACHEDLABEL [TEXTPROPS] End]
        /// </summary>
        public class LegendExceptionAggregate : ChartRecordAggregate
        {
            private LegendExceptionRecord legendException = null;
            private AttachedLabelAggregate attachedLabel = null;
            private TextPropsAggregate textProps = null;
            public LegendExceptionRecord LegendException
            {
                get { return legendException; }
            }
            public LegendExceptionAggregate(RecordStream rs, ChartRecordAggregate container)
                : base(RuleName_LEGENDEXCEPTION, container)
            {
                legendException = (LegendExceptionRecord)rs.GetNext();
                if (rs.PeekNextChartSid() == BeginRecord.sid)
                {
                    rs.GetNext();
                    attachedLabel = new AttachedLabelAggregate(rs, this);
                    if (rs.PeekNextChartSid() == TextPropsStreamRecord.sid || 
                        rs.PeekNextChartSid() == RichTextStreamRecord.sid)
                    {
                        textProps = new TextPropsAggregate(rs, this);
                    }
                    Record r = rs.GetNext();//EndRecord
                    Debug.Assert(r.GetType() == typeof(EndRecord));
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

                    WriteEndBlock(rv);
                    rv.VisitRecord(EndRecord.instance);
                }
            }
        }
    }
}

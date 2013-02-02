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
using System;
using System.Collections.Generic;
using NPOI.HSSF.Model;
using System.IO;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// CHARTSHEET = BOF CHARTSHEETCONTENT
    /// CHARTSHEETCONTENT = [WriteProtect] [SheetExt] [WebPub] *HFPicture PAGESETUP PrintSize 
    /// [HeaderFooter] [BACKGROUND] *Fbi *Fbi2 [ClrtClient] [PROTECTION] [Palette] [SXViewLink]
    /// [PivotChartBits] [SBaseRef] [MsoDrawingGroup] OBJECTS Units CHARTFOMATS SERIESDATA 
    /// *WINDOW *CUSTOMVIEW [CodeName] [CRTMLFRT] EOF
    /// </summary>
    public class ChartSheetAggregate : ChartRecordAggregate
    {
        private BOFRecord _bofRec;
        /**
         * All the records between BOF and EOF
         */
        private List<RecordBase> _recs;
        private PageSettingsBlock _psBlock;

        private ChartFormatsAggregate chartFormats;
        private SeriesDataAggregate seriesData;

        public ChartSheetAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_CHARTSHEET, container)
        {
            _bofRec = (BOFRecord)rs.GetNext();
            List<RecordBase> temp = new List<RecordBase>();
            while (rs.PeekNextClass() != typeof(EOFRecord))
            {
                Type a = rs.PeekNextClass();
                if (PageSettingsBlock.IsComponentRecord(rs.PeekNextChartSid()))
                {
                    if (_psBlock != null)
                    {
                        if (rs.PeekNextChartSid() == HeaderFooterRecord.sid)
                        {
                            // test samples: 45538_classic_Footer.xls, 45538_classic_Header.xls
                            _psBlock.AddLateHeaderFooter((HeaderFooterRecord)rs.GetNext());
                            continue;
                        }
                        throw new InvalidDataException(
                                "Found more than one PageSettingsBlock in chart sub-stream");
                    }
                    _psBlock = new PageSettingsBlock(rs);
                    temp.Add(_psBlock);
                    continue;
                }
                if (rs.PeekNextChartSid() == ChartRecord.sid)
                {
                    chartFormats = new ChartFormatsAggregate(rs, this);
                    temp.Add(chartFormats);
                    continue;
                }
                if (rs.PeekNextChartSid() == DimensionsRecord.sid)
                {
                    seriesData = new SeriesDataAggregate(rs);
                    temp.Add(seriesData);
                    continue;
                }
                temp.Add(rs.GetNext());
            }
            _recs = temp;
            Record eof = rs.GetNext(); // no need to save EOF in field
            if (!(eof is EOFRecord))
            {
                throw new InvalidOperationException("Bad chart EOF");
            }
        }
        internal int AttachLabelCount = 0;
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            if (_recs.Count == 0)
            {
                return;
            }
            rv.VisitRecord(_bofRec);

            for (int i = 0; i < _recs.Count; i++)
            {
                RecordBase rb = _recs[i];
                if (rb is RecordAggregate)
                {
                    ((RecordAggregate)rb).VisitContainedRecords(rv);
                }
                else
                {
                    rv.VisitRecord((Record)rb);
                }
            }
            rv.VisitRecord(EOFRecord.instance);
        }
    }
}

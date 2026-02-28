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
    /// FONTLIST = FrtFontList StartObject *(Font [Fbi]) EndObject
    /// </summary>
    public class FontListAggregate : ChartRecordAggregate
    {
        private FrtFontListRecord frtFontList = null;
        private ChartStartObjectRecord startObject = null;
        private Dictionary<FontRecord, FbiRecord> dicFonts = new Dictionary<FontRecord, FbiRecord>();
        private ChartEndObjectRecord endObject = null;
        public FontListAggregate(RecordStream rs, ChartRecordAggregate container)
            : base(RuleName_FONTLIST, container)
        {
            frtFontList = (FrtFontListRecord)rs.GetNext();
            startObject = (ChartStartObjectRecord)rs.GetNext();
            FontRecord f = null;
            FbiRecord fbi = null;
            while (rs.PeekNextChartSid() == FontRecord.sid)
            {
                f = (FontRecord)rs.GetNext();
                if (rs.PeekNextChartSid() == FbiRecord.sid)
                {
                    fbi = (FbiRecord)rs.GetNext();
                }
                else
                {
                    fbi = null;
                }
                dicFonts.Add(f, fbi);
            }

            endObject = (ChartEndObjectRecord)rs.GetNext();
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            WriteStartBlock(rv);//??
            rv.VisitRecord(frtFontList);
            rv.VisitRecord(startObject);
            IsInStartObject = true;
            foreach (KeyValuePair<FontRecord, FbiRecord> kv in dicFonts)
            {
                rv.VisitRecord(kv.Key);
                if (kv.Value != null)
                    rv.VisitRecord(kv.Value);
            }
            IsInStartObject = false;
            rv.VisitRecord(endObject);
        }
    }
}

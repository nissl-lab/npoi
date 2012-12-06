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
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// GELFRAME = 1*2GelFrame *Continue [PICF]
    /// PICF = Begin PicF End
    /// </summary>
    public class GelFrameAggregate : RecordAggregate
    {
        private GelFrameRecord gelFrame1;
        private GelFrameRecord gelFrame2;
        private List<ContinueRecord> continues = new List<ContinueRecord>();
        private PicFRecord picF;
        public GelFrameAggregate(RecordStream rs)
        {
            gelFrame1 = (GelFrameRecord)rs.GetNext();
            int sid = rs.PeekNextSid();
            if (sid == GelFrameRecord.sid)
            {
                gelFrame2 = (GelFrameRecord)rs.GetNext();
                sid = rs.PeekNextSid();
            }
            if (sid == ContinueRecord.sid)
            {
                while (rs.PeekNextSid() == ContinueRecord.sid)
                {
                    continues.Add((ContinueRecord)rs.GetNext());
                }
            }
            if (rs.PeekNextSid() == BeginRecord.sid)
            {
                rs.GetNext();
                picF = (PicFRecord)rs.GetNext();
                rs.GetNext();
            }
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(gelFrame1);
            if (gelFrame2 != null)
                rv.VisitRecord(gelFrame2);
            foreach (ContinueRecord cr in continues)
                rv.VisitRecord(cr);
            if (picF != null)
            {
                rv.VisitRecord(BeginRecord.instance);
                rv.VisitRecord(picF);
                rv.VisitRecord(EndRecord.instance);
            }
        }
    }
}

/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record.Aggregates
{
    using System;


    using NPOI.HSSF.Model;
    using System.Collections.Generic;
    using System.IO;

    /**
     * Manages the all the records associated with a chart sub-stream.<br/>
     * Includes the Initial {@link BOFRecord} and {@link EOFRecord}.
     *
     * @author Josh Micich
     */
    public class ChartSubstreamRecordAggregate : RecordAggregate
    {

        private BOFRecord _bofRec;
        /**
         * All the records between BOF and EOF
         */
        private List<RecordBase> _recs;
        private PageSettingsBlock _psBlock;

        public ChartSubstreamRecordAggregate(RecordStream rs)
        {
            _bofRec = (BOFRecord)rs.GetNext();
            List<RecordBase> temp = new List<RecordBase>();
            while (rs.PeekNextClass() != typeof(EOFRecord))
            {
                Type a = rs.PeekNextClass();
                if (PageSettingsBlock.IsComponentRecord(rs.PeekNextSid()))
                {
                    if (_psBlock != null)
                    {
                        if (rs.PeekNextSid() == HeaderFooterRecord.sid)
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
                temp.Add(rs.GetNext());
            }
            _recs = temp;
            Record eof = rs.GetNext(); // no need to save EOF in field
            if (!(eof is EOFRecord))
            {
                throw new InvalidOperationException("Bad chart EOF");
            }
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            if (_recs.Count==0)
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



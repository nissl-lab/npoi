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
    using NPOI.HSSF.Record;
    using System.Collections.Generic;

    /**
     * Manages the all the records associated with a 'Custom View Settings' sub-stream.<br/>
     * Includes the Initial USERSVIEWBEGIN(0x01AA) and USERSVIEWEND(0x01AB).
     * 
     * @author Josh Micich
     */
    public class CustomViewSettingsRecordAggregate : RecordAggregate
    {

        private Record _begin;
        private Record _end;
        /**
         * All the records between BOF and EOF
         */
        private List<RecordBase> _recs;
        private PageSettingsBlock _psBlock;

        public CustomViewSettingsRecordAggregate(RecordStream rs)
        {
            _begin = rs.GetNext();
            if (_begin.Sid != UserSViewBegin.sid)
            {
                throw new InvalidOperationException("Bad begin record");
            }
            List<RecordBase> temp = new List<RecordBase>();
            while (rs.PeekNextSid() != UserSViewEnd.sid)
            {
                if (PageSettingsBlock.IsComponentRecord(rs.PeekNextSid()))
                {
                    if (_psBlock != null)
                    {
                        throw new InvalidOperationException(
                                "Found more than one PageSettingsBlock in custom view Settings sub-stream");
                    }
                    _psBlock = new PageSettingsBlock(rs);
                    temp.Add(_psBlock);
                    continue;
                }
                temp.Add(rs.GetNext());
            }
            _recs = temp;
            _end = rs.GetNext(); // no need to save EOF in field
            if (_end.Sid != UserSViewEnd.sid)
            {
                throw new InvalidOperationException("Bad custom view Settings end record");
            }
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            if (_recs.Count == 0)
            {
                return;
            }
            rv.VisitRecord(_begin);
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
            rv.VisitRecord(_end);
        }

        public static bool IsBeginRecord(int sid)
        {
            return sid == UserSViewBegin.sid;
        }

        public void Append(RecordBase r)
        {
            _recs.Add(r);
        }
    }
}


using System.Collections.Generic;
using NPOI.HSSF.Model;

namespace NPOI.HSSF.Record.Aggregates
{
    internal class PLSAggregate : RecordAggregate
    {
        private static ContinueRecord[] EMPTY_CONTINUE_RECORD_ARRAY = { };
        private Record _pls;
        /**
         * holds any continue records found after the PLS record.<br/>
         * This would not be required if PLS was properly interpreted.
         * Currently, PLS is an {@link UnknownRecord} and does not automatically
         * include any trailing {@link ContinueRecord}s.
         */
        private ContinueRecord[] _plsContinues;

        public PLSAggregate(RecordStream rs)
        {
            _pls = rs.GetNext();
            if (rs.PeekNextSid() == ContinueRecord.sid)
            {
                List<ContinueRecord> temp = new List<ContinueRecord>();
                while (rs.PeekNextSid() == ContinueRecord.sid)
                {
                    temp.Add((ContinueRecord)rs.GetNext());
                }
                _plsContinues = new ContinueRecord[temp.Count];
                _plsContinues = temp.ToArray();
            }
            else
            {
                _plsContinues = EMPTY_CONTINUE_RECORD_ARRAY;
            }
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(_pls);
            for (int i = 0; i < _plsContinues.Length; i++)
            {
                rv.VisitRecord(_plsContinues[i]);
            }
        }

    }
}

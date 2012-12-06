using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Record.Chart;
using NPOI.HSSF.Model;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// DAT = Dat Begin LD End
    /// </summary>
    public class DatAggregate : RecordAggregate
    {
        private DatRecord dat = null;
        private LDAggregate ld = null;

        public DatAggregate(RecordStream rs)
        {
            dat = (DatRecord)rs.GetNext();
            rs.GetNext();
            ld = new LDAggregate(rs);
            rs.GetNext();
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(dat);
            rv.VisitRecord(BeginRecord.instance);
            ld.VisitContainedRecords(rv);
            rv.VisitRecord(EndRecord.instance);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Record.Aggregates;
using NPOI.HSSF.Record;

namespace NPOI.Tools.POIFSBrowser
{
    public class MockRecordVisitor : RecordVisitor
    {

        private List<Record> _list;
        public MockRecordVisitor()
        {
            _list = new List<Record>();
        }
        public void VisitRecord(Record r)
        {
            _list.Add(r);
        }
        public List<Record> Records
        {
            get { return _list; }
        }
    }
}

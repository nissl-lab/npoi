using System;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using NPOI.HSSF.Record;
using NPOI.HSSF.Util;


namespace NPOI.Tools.POIFSBrowser
{
    internal class CellValueRecordTreeNode : AbstractRecordTreeNode
    {
        public CellValueRecordTreeNode(CellValueRecordInterface record)
        {
            this.Record = record;
            this.Text = record.GetType().Name;
            this.ImageKey = "Binary";
        }

        public override bool HasBinary
        {
            get { return false; }
        }
        public override byte[] GetBytes()
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using NPOI.HSSF.Record;

namespace NPOI.Tools.POIFSBrowser
{
    internal class UnicodeStringTreeNode : AbstractRecordTreeNode
    {
        public UnicodeStringTreeNode(UnicodeString record)
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

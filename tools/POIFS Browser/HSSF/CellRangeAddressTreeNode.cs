using System;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using NPOI.HSSF.Record;
using NPOI.HSSF.Util;
using NPOI.Util;

namespace NPOI.Tools.POIFSBrowser
{
    internal class CellRangeAddressTreeNode : AbstractRecordTreeNode
    {
        public NPOI.SS.Util.CellRangeAddress Record { get; private set; }

        public CellRangeAddressTreeNode(NPOI.SS.Util.CellRangeAddress record)
        {
            this.Record = record;
            this.Text = record.GetType().Name;
            this.ImageKey = "Binary";
        }

         public override ListViewItem[] GetPropertyList()
         {
             ArrayList tmplist = new ArrayList();

             tmplist.Add(new ListViewItem(
                     new string[] { "", "FirstRow", this.Record.FirstRow.GetType().Name, this.Record.FirstRow.ToString() }));
             tmplist.Add(new ListViewItem(
        new string[] { "", "FirstColumn", this.Record.FirstColumn.GetType().Name, this.Record.FirstColumn.ToString() }));
             tmplist.Add(new ListViewItem(
        new string[] { "", "LastRow", this.Record.LastRow.GetType().Name, this.Record.LastRow.ToString() }));

             tmplist.Add(new ListViewItem(
        new string[] { "", "LastColumn", this.Record.LastColumn.GetType().Name, this.Record.LastColumn.ToString() }));


             return (ListViewItem[])tmplist.ToArray(typeof(ListViewItem));
         }

         public override byte[] GetBytes()
         {

             byte[] bytes = new byte[NPOI.SS.Util.CellRangeAddress.ENCODED_SIZE];
             ILittleEndianOutput leo = new LittleEndianByteArrayOutputStream(bytes, 0);
             this.Record.Serialize(leo);
             return bytes;
         }

         public override bool HasBinary
         {
             get { return true; }
         }
    }
}

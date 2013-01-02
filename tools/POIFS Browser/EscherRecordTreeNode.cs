using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

using NPOI.DDF;

namespace NPOI.Tools.POIFSBrowser
{
    public class EscherRecordTreeNode : AbstractRecordTreeNode
    {
        public EscherRecordTreeNode(EscherRecord record)
        {
            this.Record = record;
            this.Text = record.GetType().Name;
            if(record.ChildRecords.Count>0)
                this.ImageKey = "Folder";
            else
                this.ImageKey = "Binary";

            if (record is AbstractEscherOptRecord)
            {
                AbstractEscherOptRecord rec = ((AbstractEscherOptRecord)record);
                foreach (EscherProperty ep in rec.EscherProperties )
                {
                    EscherPropertyTreeNode cnode = new EscherPropertyTreeNode(ep);
                    this.Nodes.Add(cnode);
                }
            }
            else
            {
                GetChildren(this);
            }
        }
        
        private void GetChildren(EscherRecordTreeNode node)
        {
            EscherRecord record=((EscherRecord)node.Record);
            for (int i = 0; i < record.ChildRecords.Count; i++)
            {
                EscherRecordTreeNode cnode=new EscherRecordTreeNode((EscherRecord)record.ChildRecords[i]);
                this.Nodes.Add(cnode);
            }
        }

        public override byte[] GetBytes()
        {
             EscherRecord record=(EscherRecord)this.Record;
             byte[] bytes=new byte[record.RecordSize];
             record.Serialize(0, bytes);
             return bytes;
        }

        public override bool HasBinary
        {
            get { return true; }
        }
    }
}

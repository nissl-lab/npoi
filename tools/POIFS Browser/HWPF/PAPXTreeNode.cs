using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HWPF.Model;

namespace NPOI.Tools.POIFSBrowser.HWPF
{
    public class PAPXTreeNode:AbstractHWPFTreeNode
    {
        public PAPXTreeNode(PAPX papx)
        {
            this.Record = papx;
        }

        public override byte[] GetBytes()
        {
            throw new NotImplementedException();
        }
    }
}

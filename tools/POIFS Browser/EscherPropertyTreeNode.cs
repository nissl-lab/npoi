using System;
using System.Collections.Generic;
using System.Text;
using NPOI.DDF;

namespace NPOI.Tools.POIFSBrowser
{
    public class EscherPropertyTreeNode:AbstractRecordTreeNode
    {
        

        public EscherPropertyTreeNode(EscherProperty ep)
        {
            this.Record = ep;
            this.Text = ep.Name;
             this.ImageKey = "Binary";
        }

        public override bool HasBinary
        {
            get { return true; }
        }

        public override byte[] GetBytes()
        {

            if (this.Record is EscherSimpleProperty)
            {
                EscherSimpleProperty esp = (EscherSimpleProperty)this.Record;
                byte[] data = new byte[esp.PropertySize];
                esp.SerializeSimplePart(data, 0);
                return data;
            }
            else if (this.Record is EscherComplexProperty)
            {
                EscherComplexProperty esp = (EscherComplexProperty)this.Record;
                byte[] data = new byte[esp.PropertySize];
                esp.SerializeComplexPart(data, 0);
                return data;
            }
            return null;
            
        }
    }
}

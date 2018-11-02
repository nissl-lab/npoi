using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class FtrDocument
    {

        CT_Ftr ftr = null;
        public FtrDocument()
        {
            ftr = new CT_Ftr();
        }
        public static FtrDocument Parse(XmlDocument doc, XmlNamespaceManager namespaceMgr)
        {
            CT_Ftr obj = CT_Ftr.Parse(doc.DocumentElement, namespaceMgr);
            return new FtrDocument(obj);
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                ftr.Write(sw);
            }
        }

        public FtrDocument(CT_Ftr ftr)
        {
            this.ftr = ftr;
        }
        public CT_Ftr Ftr
        {
            get
            {
                return this.ftr;
            }
        }
        public void SetFtr(CT_Ftr ftr)
        {
            this.ftr = ftr;
        }
    }
}

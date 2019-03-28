using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class HdrDocument
    {

        CT_Hdr hdr = null;
        public HdrDocument()
        {
            hdr = new CT_Hdr();
        }
        public static HdrDocument Parse(XmlDocument doc, XmlNamespaceManager namespaceMgr)
        {
            CT_Hdr obj = CT_Hdr.Parse(doc.DocumentElement, namespaceMgr);
            return new HdrDocument(obj);
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                hdr.Write(sw);
            }
        }
        public HdrDocument(CT_Hdr hdr)
        {
            this.hdr = hdr;
        }
        public CT_Hdr Hdr
        {
            get
            {
                return this.hdr;
            }
        }

        public void SetHdr(CT_Hdr hdr)
        {
            this.hdr = hdr;
        }
    }
}

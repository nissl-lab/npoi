using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet.Document
{
    public class ExternalLinkDocument
    {
        CT_ExternalLink link = null;

        public ExternalLinkDocument()
        {
        }
        public ExternalLinkDocument(CT_ExternalLink link)
        {
            this.link = link;
        }
        public static ExternalLinkDocument Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceMgr)
        {
            CT_ExternalLink obj = CT_ExternalLink.Parse(xmldoc.DocumentElement, namespaceMgr);
            return new ExternalLinkDocument(obj);
        }

        public CT_ExternalLink ExternalLink
        {
            get
            {
                return link;
            }
            set
            {
                this.link = value;
            }
        }
        public void Save(Stream stream)
        {
            using (StreamWriter sw1 = new StreamWriter(stream))
            {
                this.link.Write(sw1);
            }
        }
    }
}

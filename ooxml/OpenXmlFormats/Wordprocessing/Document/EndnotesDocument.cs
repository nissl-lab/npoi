using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class EndnotesDocument
    {
        CT_Endnotes endnotes = null;
        public EndnotesDocument()
        {

        }
        public static EndnotesDocument Parse(XmlDocument doc, XmlNamespaceManager namespaceMgr)
        {
            CT_Endnotes obj = CT_Endnotes.Parse(doc.DocumentElement, namespaceMgr);
            return new EndnotesDocument(obj);
        }
        public EndnotesDocument(CT_Endnotes endnotes)
        {
            this.endnotes = endnotes;
        }
        public CT_Endnotes Endnotes
        {
            get
            {
                return this.endnotes;
            }
        }
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                this.endnotes.Write(sw);
            }
        }
    }
    
}

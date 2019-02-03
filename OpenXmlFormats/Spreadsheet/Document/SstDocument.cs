using NPOI.OpenXml4Net.Util;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class SstDocument
    {
        CT_Sst sst = null;

        public SstDocument()
        {         
        }
        public SstDocument(CT_Sst sst)
        {
            this.sst = sst;
        }

        public void AddNewSst()
        {
            this.sst = new CT_Sst();
        }

        public CT_Sst GetSst()
        {
            return this.sst;
        }
        public static SstDocument Parse(XmlDocument xml, XmlNamespaceManager namespaceManager)
        {

            try
            {
                SstDocument sstDoc=new SstDocument();
                sstDoc.AddNewSst();
                CT_Sst sst = sstDoc.GetSst();
                sst.count = XmlHelper.ReadInt(xml.DocumentElement.Attributes["count"]);
                sst.uniqueCount = XmlHelper.ReadInt(xml.DocumentElement.Attributes["uniqueCount"]);

                XmlNodeList nl = xml.SelectNodes("//d:sst/d:si", namespaceManager);
                if (nl != null)
                {
                    foreach (XmlNode node in nl)
                    {
                        CT_Rst rst = CT_Rst.Parse(node, namespaceManager);
                        sstDoc.sst.si.Add(rst);
                    }
                }
                return sstDoc;
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }        

        public void Save(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream, Encoding.UTF8);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{1}\">", this.GetSst().count, this.GetSst().uniqueCount);
            foreach (CT_Rst ssi in this.GetSst().si)
            {
                ssi.Write(sw, "si");
            }
            sw.Write("</sst>");
            sw.Flush();
        }

    }
}

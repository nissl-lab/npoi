using NPOI.OpenXml4Net.Util;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CalcChainDocument
    {
        CT_CalcChain calcChain;

        public CalcChainDocument()
        {
            this.calcChain = new CT_CalcChain();
        }
        internal CalcChainDocument(CT_CalcChain calcChain)
        {
            this.calcChain = calcChain;
        }

        public CT_CalcChain GetCalcChain()
        {
            return calcChain;
        }

        public void SetCalcChain(CT_CalcChain calcchain)
        {
            this.calcChain = calcchain;
        }

        public static CalcChainDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager NameSpaceManager)
        {
            CalcChainDocument calcChainDoc = new CalcChainDocument();
            foreach (XmlElement node in xmlDoc.SelectNodes("//d:c", NameSpaceManager))
            {
                CT_CalcCell cc = new CT_CalcCell();
                if (node.GetAttributeNode("i")!= null)
                {
                    cc.i = XmlHelper.ReadInt(node.GetAttributeNode("i"));
                    cc.iSpecified = true;
                }
                cc.r = node.GetAttribute("r");
                cc.t = XmlHelper.ReadBool(node.GetAttributeNode("t"));
                cc.s = XmlHelper.ReadBool(node.GetAttributeNode("s"));
                cc.l = XmlHelper.ReadBool(node.GetAttributeNode("l"));
                calcChainDoc.calcChain.AddC(cc);
            }
            return calcChainDoc;
        }

        public void Save(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
            sw.Write("<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            foreach (CT_CalcCell cc in calcChain.c)
            {
                sw.Write("<c");
                sw.Write(" r=\""+cc.r+"\"");
                if(cc.i>0)
                    sw.Write(" i=\"" + cc.i + "\"");
                if(cc.s)
                    sw.Write(" s=\"" + (cc.s?1:0) + "\"");
                if (cc.t)
                    sw.Write(" t=\"" + (cc.t?1:0) + "\"");
                if (cc.l)
                    sw.Write(" l=\"" + (cc.l?1:0) + "\"");
                sw.Write("/>");

            }
            sw.Write("</calcChain>");
            sw.Flush();
        }

    }
}

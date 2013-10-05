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
            var calcChainDoc = new CalcChainDocument();
            foreach (XmlElement node in xmlDoc.SelectNodes("//d:c", NameSpaceManager))
            {
                CT_CalcCell cc = new CT_CalcCell();
                if (node.GetAttributeNode("i") != null)
                {
                    cc.i = Int32.Parse(node.GetAttribute("i"));
                    cc.iSpecified = true;
                }
                cc.r = node.GetAttribute("r");
                if (node.GetAttributeNode("t") != null)
                {
                    string value = node.GetAttribute("t");
                    if (value == "1" || value.ToLower() == "true")
                    {
                        cc.t = true;
                    }
                    else
                    {
                        cc.t = false;
                    }
                }
                if (node.GetAttributeNode("s") != null)
                {
                    string value = node.GetAttribute("s");
                    if (value == "1" || value.ToLower() == "true")
                    {
                        cc.s = true;
                    }
                    else
                    {
                        cc.s = false;
                    }
                }
                if (node.GetAttributeNode("l") != null)
                {
                    string value = node.GetAttribute("l");
                    if (value == "1" || value.ToLower() == "true")
                    {
                        cc.l = true;
                    }
                    else
                    {
                        cc.l = false;
                    }
                }

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

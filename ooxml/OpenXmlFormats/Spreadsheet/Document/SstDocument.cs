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

                XmlNodeList nl = xml.SelectNodes("//d:sst/d:si", namespaceManager);
                if (nl != null)
                {
                    foreach (XmlNode node in nl)
                    {
                        XmlNode n = node.SelectSingleNode("d:t", namespaceManager);
                        CT_Rst rst = new CT_Rst();
                        rst.XmlText = node.InnerXml;
                        if (n != null)
                        {
                            rst.t = n.InnerText;
                        }
                        else
                        {
                            XmlNodeList tNodes = node.SelectNodes(".//d:t", namespaceManager);
                            if (tNodes != null)
                            {
                                rst.r = new System.Collections.Generic.List<CT_RElt>();

                                foreach (XmlNode tNode in tNodes)
                                {
                                    CT_RElt relt = new CT_RElt();
                                    relt.t = tNode.InnerText;
                                    rst.r.Add(relt);
                                }
                            }
                        }
                        sstDoc.sst.si.Add(rst);
                    }
                }
                return sstDoc;
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="sw"></param>
        /// <param name="t"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StreamWriter sw, string t)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
            {
                Match match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
                int indexAdd = 0;
                while (match.Success)
                {
                    t = t.Insert(match.Index + indexAdd, "_x005F");
                    indexAdd += 6;
                    match = match.NextMatch();
                }
            }
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] <= 0x1f && t[i] != '\t' && t[i] != '\n' && t[i] != '\r') //Not Tab, CR or LF
                {
                    sw.Write("_x00{0}_", (t[i] < 0xa ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else
                {
                    sw.Write(t[i]);
                }
            }

        }
        internal static string ExcelDecodeString(string t)
        {
            Match match = Regex.Match(t, "(_x005F|_x[0-9A-F]{4,4}_)");
            if (!match.Success) return t;

            bool useNextValue = false;
            StringBuilder ret = new StringBuilder();
            int prevIndex = 0;
            while (match.Success)
            {
                if (prevIndex < match.Index) ret.Append(t.Substring(prevIndex, match.Index - prevIndex));
                if (!useNextValue && match.Value == "_x005F")
                {
                    useNextValue = true;
                }
                else
                {
                    if (useNextValue)
                    {
                        ret.Append(match.Value);
                        useNextValue = false;
                    }
                    else
                    {
                        ret.Append((char)int.Parse(match.Value.Substring(2, 4)));
                    }
                }
                prevIndex = match.Index + match.Length;
                match = match.NextMatch();
            }
            ret.Append(t.Substring(prevIndex, t.Length - prevIndex));
            return ret.ToString();
        }
        public void Save(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{1}\">", this.GetSst().count, this.GetSst().uniqueCount);
            foreach (CT_Rst ssi in this.GetSst().si)
            {
                string t = ssi.t;
                if (ssi.XmlText != null)
                {
                    sw.Write("<si>");
                    ExcelEncodeString(sw, ssi.XmlText);
                    sw.Write("</si>");
                }
                else
                {
                    if (t.Length > 0 && (t[0] == ' ' || t[t.Length - 1] == ' ' || t.Contains("  ") || t.Contains("\t")))
                    {
                        sw.Write("<si><t xml:space=\"preserve\">");
                    }
                    else
                    {
                        sw.Write("<si><t>");
                    }
                    ExcelEncodeString(sw,NPOI.OpenXml4Net.Util.XmlHelper.EncodeXml(t));
                    sw.Write("</t></si>");
                }
            }
            sw.Write("</sst>");
            sw.Flush();
        }

    }
}

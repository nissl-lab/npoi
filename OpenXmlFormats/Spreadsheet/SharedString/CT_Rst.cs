using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Rst
    {

        private string tField = null; // optional field -> initialize as null so that it is not serialized by default.

        private List<CT_RElt> rField = null; // optional field 

        private List<CT_PhoneticRun> rPhField = null; // optional field 

        private CT_PhoneticPr phoneticPrField = null; // optional field 

        public void Set(CT_Rst o)
        {
            this.tField = o.tField;
            this.rField = o.rField;
            this.rPhField = o.rPhField;
            this.phoneticPrField = o.phoneticPrField;
        }
      



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}>", nodeName));
            if (this.t != null)
            {
                //TODO: diff has-space case and no-space case
                 sw.Write(string.Format("<t xml:space=\"preserve\">{0}</t>", 
                      XmlHelper.ExcelEncodeString(XmlHelper.EncodeXml(t))));
            }
            if (this.r != null)
            {
                foreach (CT_RElt x in this.r)
                {
                    x.Write(sw, "r");
                }
            }
            if (this.rPh != null)
            {
                foreach (CT_PhoneticRun x in this.rPh)
                {
                    x.Write(sw, "rPh");
                }
            }
            if (this.phoneticPr != null)
                this.phoneticPr.Write(sw, "phoneticPr");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        #region t
        public bool IsSetT()
        {
            return this.tField != null;
        }
        public void unsetT()
        {
            this.tField = null;
        }
        [XmlElement("t", DataType = "string")]
        public string t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }
        #endregion t

        #region r
        /// <summary>
        /// Rich Text Run
        /// </summary>
        [XmlElement("r")]
        public List<CT_RElt> r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }
        private string xmltext;
        [XmlIgnore]
        public string XmlText
        {
            get {
                StringBuilder sb = new StringBuilder();
                using (StringWriter sw = new StringWriter(sb))
                {
                    if (rField != null && rField.Count > 0)
                    {

                        foreach (CT_RElt r in rField)
                        {
                            sw.Write("<r>");
                            if (r.rPr != null)
                            {
                                sw.Write("<rPr>");
                                if (r.rPr.b != null && r.rPr.b.val)
                                {
                                    sw.Write("<b/>");
                                }
                                if (r.rPr.i != null && r.rPr.i.val)
                                {
                                    sw.Write("<i/>");
                                }
                                if (r.rPr.u != null)
                                {
                                    sw.Write("<u val=\"" + r.rPr.u.val + "\"/>");
                                }
                                if (r.rPr.color != null && r.rPr.color.theme > 0)
                                {
                                    sw.Write("<color theme=\"" + r.rPr.color.theme + "\"/>");
                                }
                                if (r.rPr.color != null && r.rPr.color.rgbSpecified)
                                {
                                    sw.Write("<color rgb=\"" + BitConverter.ToString(r.rPr.color.rgb).Replace("-", string.Empty) + "\"/>");
                                }
                                if (r.rPr.rFont != null)
                                {
                                    sw.Write("<rFont val=\"" + r.rPr.rFont.val + "\"/>");
                                }
                                if (r.rPr.family != null)
                                {
                                    sw.Write("<family val=\"" + r.rPr.family.val + "\"/>");
                                }
                                if (r.rPr.charset != null)
                                {
                                    sw.Write("<charset val=\"" + r.rPr.charset.val + "\"/>");
                                }
                                if (r.rPr.scheme != null)
                                {
                                    sw.Write("<scheme val=\"" + r.rPr.scheme.val + "\"/>");
                                }
                                if (r.rPr.sz != null)
                                {
                                    sw.Write("<sz val=\"" + r.rPr.sz.val + "\"/>");
                                }
                                if (r.rPr.vertAlign != null)
                                {
                                    sw.Write("<vertAlign val=\"" + r.rPr.vertAlign.val + "\"/>");
                                }
                                sw.Write("</rPr>");
                            }
                            if (r.t != null)
                            {
                                sw.Write("<t");
                                if(r.t.IndexOf(' ')>=0)
                                    sw.Write(" xml:space=\"preserve\"");
                                sw.Write(">");
                                sw.Write(XmlHelper.EncodeXml(r.t));
                                sw.Write("</t>");
                            }
                            sw.Write("</r>");
                        }
                    }

                    if (this.t != null)
                    {
                        sw.Write("<t");
                        if (this.t.IndexOf(' ') >= 0)
                            sw.Write(" xml:space=\"preserve\"");
                        sw.Write(">");
                        sw.Write(XmlHelper.EncodeXml(this.t));
                        sw.Write("</t>");
                    }
                    xmltext = sb.ToString();
                }
                return xmltext; 
            }
            set { xmltext = value; }
        }
        public CT_RElt AddNewR()
        {
            if (null == rField) { rField = new List<CT_RElt>(); }
            CT_RElt r = new CT_RElt();
            this.rField.Add(r);
            return r;
        }
        public int sizeOfRArray()
        {
            return (null == rField) ? 0 : r.Count;
        }
        public CT_RElt GetRArray(int index)
        {
            return (null == rField) ? null : this.rField[index];
        }
        #endregion r

        /// <summary>
        /// Phonetic Run
        /// </summary>
        [XmlElement("rPh")]
        public List<CT_PhoneticRun> rPh
        {
            get
            {
                return this.rPhField;
            }
            set
            {
                this.rPhField = value;
            }
        }
        /// <summary>
        /// Phonetic Properties
        /// </summary>
        [XmlElement("phoneticPr")]
        public CT_PhoneticPr phoneticPr
        {
            get
            {
                return this.phoneticPrField;
            }
            set
            {
                this.phoneticPrField = value;
            }
        }


        public static CT_Rst Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_Rst ctObj = new CT_Rst();
            ctObj.r = new List<CT_RElt>();
            ctObj.rPh = new List<CT_PhoneticRun>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "phoneticPr")
                    ctObj.phoneticPr = CT_PhoneticPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "r")
                    ctObj.r.Add(CT_RElt.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "rPh")
                    ctObj.rPh.Add(CT_PhoneticRun.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "t")
                    ctObj.t = childNode.InnerText.Replace("\r", "");
            }
            return ctObj;
        }

        public int SizeOfRArray()
        {
            return r.Count;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;

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

                if (rField ==null || rField.Count==0)
                    this.xmltext = "<r><t xml:space=\"preserve\">"+value+"</t></r>";
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
        [XmlText]
        public string XmlText
        {
            get {
                if (rField != null && rField.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    using( StringWriter sw = new StringWriter(sb))
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
                                    sw.Write("<u val=\""+ r.rPr.u.val +"\"/>");
                                }
                                if (r.rPr.color != null && r.rPr.color.theme > 0)
                                {
                                    sw.Write("<color theme=\"" + r.rPr.color.theme + "\"/>");
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
                                sw.Write("<t xml:space=\"preserve\">");
                                sw.Write(r.t);
                                sw.Write("</t>");
                            }
                            sw.Write("</r>");
                        }
                        xmltext = sb.ToString();
                    }
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


        public static CT_Rst Parse(XmlNode xmlNode, XmlNamespaceManager namespaceManager)
        {
            CT_Rst rst = new CT_Rst();
            rst.r = new List<CT_RElt>();
            var rNodes = xmlNode.SelectNodes("d:r", namespaceManager);
            foreach (XmlNode rNode in rNodes)
            {
                CT_RElt relt = rst.AddNewR();
                var rPrNode = rNode.SelectSingleNode("d:rPr", namespaceManager);
                if (rPrNode != null)
                {
                    CT_RPrElt rprelt = relt.AddNewRPr();
                    foreach (XmlNode childNode in rPrNode.ChildNodes)
                    {
                        switch (childNode.Name)
                        { 
                            case "b":
                                CT_BooleanProperty bprop= rprelt.AddNewB();
                                bprop.val = true;
                                break;
                            case "i":
                                CT_BooleanProperty iprop = rprelt.AddNewI();
                                iprop.val = true;
                                break;
                            case "u":
                                CT_UnderlineProperty uProp = rprelt.AddNewU();
                                uProp.val = (ST_UnderlineValues)Enum.Parse(typeof(ST_UnderlineValues), childNode.Attributes["val"].Value);
                                break;
                            case "color":
                                CT_Color color = rprelt.AddNewColor();
                                if(childNode.Attributes["theme"]!=null)
                                    color.theme = uint.Parse(childNode.Attributes["theme"].Value);
                                if(childNode.Attributes["auto"]!=null)
                                    color.auto = childNode.Attributes["auto"].Value=="1"?true:false;
                                if(childNode.Attributes["indexed"]!=null)
                                    color.indexed = uint.Parse(childNode.Attributes["indexed"].Value);
                                if(childNode.Attributes["tint"]!=null)
                                    color.tint = Double.Parse(childNode.Attributes["tint"].Value);
                                break;
                            case "rFont":
                                CT_FontName fontname = rprelt.AddNewRFont();
                                fontname.val = childNode.Attributes["val"].Value;
                                break;
                            case "family":
                                CT_IntProperty familyProp = rprelt.AddNewFamily();
                                familyProp.val = Int32.Parse(childNode.Attributes["val"].Value);
                                break;
                            case "charset":
                                CT_IntProperty charsetProp = rprelt.AddNewCharset();
                                charsetProp.val = Int32.Parse(childNode.Attributes["val"].Value);
                                break;
                            case "scheme":
                                CT_FontScheme schemeProp = rprelt.AddNewScheme();
                                schemeProp.val = (ST_FontScheme)Enum.Parse(typeof(ST_FontScheme), childNode.Attributes["val"].Value);
                                break;
                            case "sz":
                                CT_FontSize szProp = rprelt.AddNewSz();
                                szProp.val = Int32.Parse(childNode.Attributes["val"].Value);
                                break;
                            case "vertAlign":
                                CT_VerticalAlignFontProperty vertAlignProp = rprelt.AddNewVertAlign();
                                vertAlignProp.val = (ST_VerticalAlignRun)Enum.Parse(typeof(ST_VerticalAlignRun), childNode.Attributes["val"].Value);
                                break;
                        }
                    }
                }
                var tNode = rNode.SelectSingleNode("d:t", namespaceManager);
                relt.t = tNode.InnerText.Replace("\r","");
            }
            return rst;
        }
    }
}

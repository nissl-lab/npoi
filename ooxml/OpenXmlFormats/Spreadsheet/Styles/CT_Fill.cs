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
    public class CT_Fill
    {
        private CT_PatternFill patternFillField = null;
        private CT_GradientFill gradientFillField = null;

        [XmlElement]
        public CT_PatternFill patternFill
        {
            get { return this.patternFillField; }
            set { this.patternFillField = value; }
        }
        [XmlElement]
        public CT_GradientFill gradientFill
        {
            get { return this.gradientFillField; }
            set { this.gradientFillField = value; }
        }

        public CT_PatternFill GetPatternFill()
        {
            return this.patternFillField;
        }

        public CT_PatternFill AddNewPatternFill()
        {
            this.patternFillField = new CT_PatternFill();
            return GetPatternFill();
        }
        public bool IsSetPatternFill()
        {
            return this.patternFillField != null;
        }
        public CT_Fill Copy()
        {
            CT_Fill obj = new CT_Fill();
            obj.patternFillField = this.patternFillField;
            obj.gradientFillField = this.gradientFillField;
            return obj;
        }
        public static CT_Fill Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Fill ctObj = new CT_Fill();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "patternFill")
                    ctObj.patternFill = CT_PatternFill.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gradientFill")
                    ctObj.gradientFill = CT_GradientFill.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.patternFill != null)
                this.patternFill.Write(sw, "patternFill");
            if (this.gradientFill != null)
                this.gradientFill.Write(sw, "gradientFill");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fill));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        public override string ToString()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                StreamWriter sw = new StreamWriter(ms);
                this.Write(sw, "fill");
                sw.Flush();
                ms.Position = 0;
                StreamReader sr = new StreamReader(ms);
                string result = sr.ReadToEnd();
               return result;
            }
        }
    }

}

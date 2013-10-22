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
            return obj;
        }
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fill));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        public override string ToString()
        {
            StringWriter stringWriter = new StringWriter();
            serializer.Serialize(stringWriter, this, namespaces);
            return stringWriter.ToString();
        }
    }

}

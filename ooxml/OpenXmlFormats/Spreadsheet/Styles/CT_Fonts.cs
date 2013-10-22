using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    //[Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "fonts", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Fonts
    {

        private List<CT_Font> fontField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_Fonts()
        {
            this.fontField = new List<CT_Font>();
        }

        public void SetFontArray(CT_Font[] array)
        {
            if (array != null)
                fontField = new List<CT_Font>(array);
            else
                fontField.Clear();
        }
        [XmlElement]
        public List<CT_Font> font
        {
            get
            {
                return this.fontField;
            }
            set
            {
                this.fontField = value;
            }
        }
        [XmlAttribute]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [XmlIgnore]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fonts));
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

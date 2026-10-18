using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Presentation
{
    /// <summary>
    /// A partial slide layout model retaining unmodeled PresentationML content.
    /// </summary>
    [XmlRoot("sldLayout", Namespace = CT_Presentation.Namespace)]
    public class CT_SlideLayout
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlAttribute]
        public string type { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlAnyElement("cSld", Namespace = CT_Presentation.Namespace, Order = 0)]
        public XmlElement cSld { get; set; }

        [XmlAnyElement(Order = 1)]
        public XmlElement[] AnyElements { get; set; }
    }
}

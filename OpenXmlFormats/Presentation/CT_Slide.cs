using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Presentation
{
    /// <summary>
    /// A partial slide model retaining drawing XML for surgical shape edits and unsupported content.
    /// </summary>
    [XmlRoot("sld", Namespace = CT_Presentation.Namespace)]
    public class CT_Slide
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlAnyElement("cSld", Namespace = CT_Presentation.Namespace, Order = 0)]
        public XmlElement cSld { get; set; }

        [XmlAnyElement(Order = 1)]
        public XmlElement[] AnyElements { get; set; }
    }
}

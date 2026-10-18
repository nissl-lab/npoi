using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Presentation
{
    /// <summary>
    /// A partial slide master model retaining unmodeled PresentationML content.
    /// </summary>
    [XmlRoot("sldMaster", Namespace = CT_Presentation.Namespace)]
    public class CT_SlideMaster
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlAnyElement("cSld", Namespace = CT_Presentation.Namespace, Order = 0)]
        public XmlElement cSld { get; set; }

        [XmlAnyElement("clrMap", Namespace = CT_Presentation.Namespace, Order = 1)]
        public XmlElement clrMap { get; set; }

        [XmlAnyElement("sldLayoutIdLst", Namespace = CT_Presentation.Namespace, Order = 2)]
        public XmlElement sldLayoutIdLst { get; set; }

        [XmlAnyElement(Order = 3)]
        public XmlElement[] AnyElements { get; set; }
    }
}

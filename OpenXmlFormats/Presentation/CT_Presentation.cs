using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Presentation
{
    /// <summary>
    /// The initial PresentationML presentation model. Unmodeled content is retained as XML.
    /// </summary>
    [XmlRoot("presentation", Namespace = Namespace)]
    public class CT_Presentation
    {
        public const string Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        public const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlAnyElement("sldMasterIdLst", Namespace = Namespace, Order = 0)]
        public XmlElement sldMasterIdLst { get; set; }

        [XmlAnyElement("notesMasterIdLst", Namespace = Namespace, Order = 1)]
        public XmlElement notesMasterIdLst { get; set; }

        [XmlAnyElement("handoutMasterIdLst", Namespace = Namespace, Order = 2)]
        public XmlElement handoutMasterIdLst { get; set; }

        [XmlElement("sldIdLst", Namespace = Namespace, Order = 3)]
        public CT_SlideIdList sldIdLst { get; set; }

        [XmlAnyElement("sldSz", Namespace = Namespace, Order = 4)]
        public XmlElement sldSz { get; set; }

        [XmlAnyElement("notesSz", Namespace = Namespace, Order = 5)]
        public XmlElement notesSz { get; set; }

        [XmlAnyElement(Order = 6)]
        public XmlElement[] AnyElements { get; set; }
    }

    /// <summary>
    /// The ordered slide references in a presentation.
    /// </summary>
    [XmlType(Namespace = CT_Presentation.Namespace)]
    public class CT_SlideIdList
    {
        [XmlElement("sldId")]
        public List<CT_SlideIdListEntry> sldId { get; set; } = new List<CT_SlideIdListEntry>();

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }
    }

    /// <summary>
    /// A slide identifier and its package relationship identifier.
    /// </summary>
    [XmlType(Namespace = CT_Presentation.Namespace)]
    public class CT_SlideIdListEntry
    {
        [XmlAttribute]
        public uint id { get; set; }

        [XmlAttribute("id", Namespace = CT_Presentation.RelationshipsNamespace)]
        public string relationshipId { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlAnyElement]
        public XmlElement[] AnyElements { get; set; }
    }
}

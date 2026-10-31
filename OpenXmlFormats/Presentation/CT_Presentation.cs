using System;
using System.Collections.Generic;
using System.Linq;
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
        private int[] elementOrder = Array.Empty<int>();

        public const string Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        public const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlIgnore]
        public XmlElement sldMasterIdLst { get; set; }

        [XmlIgnore]
        public XmlElement notesMasterIdLst { get; set; }

        [XmlIgnore]
        public XmlElement handoutMasterIdLst { get; set; }

        [XmlIgnore]
        public CT_SlideIdList sldIdLst { get; set; }

        [XmlIgnore]
        public XmlElement sldSz { get; set; }

        [XmlIgnore]
        public XmlElement notesSz { get; set; }

        [XmlIgnore]
        public XmlElement[] AnyElements { get; set; }

        /// <summary>
        /// Gets or sets modeled and opaque children in their original XML order.
        /// </summary>
        [XmlElement("sldIdLst", typeof(CT_SlideIdList), Namespace = Namespace)]
        [XmlAnyElement]
        public object[] Content
        {
            get => PresentationXmlContent.Merge(elementOrder,
                new object[] { sldMasterIdLst }, new object[] { notesMasterIdLst },
                new object[] { handoutMasterIdLst }, new object[] { sldIdLst },
                new object[] { sldSz }, new object[] { notesSz }, AnyElements);
            set
            {
                var content = value ?? Array.Empty<object>();
                elementOrder = content.Select(GetElementGroup).ToArray();
                sldMasterIdLst = content.Where(item => GetElementGroup(item) == 0).OfType<XmlElement>().FirstOrDefault();
                notesMasterIdLst = content.Where(item => GetElementGroup(item) == 1).OfType<XmlElement>().FirstOrDefault();
                handoutMasterIdLst = content.Where(item => GetElementGroup(item) == 2).OfType<XmlElement>().FirstOrDefault();
                sldIdLst = content.OfType<CT_SlideIdList>().FirstOrDefault();
                sldSz = content.Where(item => GetElementGroup(item) == 4).OfType<XmlElement>().FirstOrDefault();
                notesSz = content.Where(item => GetElementGroup(item) == 5).OfType<XmlElement>().FirstOrDefault();
                AnyElements = content.Where(item => GetElementGroup(item) == 6).OfType<XmlElement>().ToArray();
            }
        }

        private static int GetElementGroup(object item)
        {
            if (item is CT_SlideIdList)
            {
                return 3;
            }
            if (item is XmlElement element && element.NamespaceURI == Namespace)
            {
                switch (element.LocalName)
                {
                    case "sldMasterIdLst": return 0;
                    case "notesMasterIdLst": return 1;
                    case "handoutMasterIdLst": return 2;
                    case "sldSz": return 4;
                    case "notesSz": return 5;
                }
            }
            return 6;
        }
    }

    /// <summary>
    /// The ordered slide references in a presentation.
    /// </summary>
    [XmlType(Namespace = CT_Presentation.Namespace)]
    public class CT_SlideIdList
    {
        private int[] elementOrder = Array.Empty<int>();

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlIgnore]
        public List<CT_SlideIdListEntry> sldId { get; set; } = new List<CT_SlideIdListEntry>();

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlIgnore]
        public XmlElement[] AnyElements { get; set; }

        /// <summary>
        /// Gets or sets slide references and opaque compatibility markup in XML order.
        /// </summary>
        [XmlElement("sldId", typeof(CT_SlideIdListEntry), Namespace = CT_Presentation.Namespace)]
        [XmlAnyElement]
        public object[] Content
        {
            get => PresentationXmlContent.Merge(elementOrder, sldId?.ToArray(), AnyElements);
            set
            {
                var content = value ?? Array.Empty<object>();
                elementOrder = content.Select(item => item is CT_SlideIdListEntry ? 0 : 1).ToArray();
                sldId = content.OfType<CT_SlideIdListEntry>().ToList();
                AnyElements = content.OfType<XmlElement>().ToArray();
            }
        }
    }

    /// <summary>
    /// A slide identifier and its package relationship identifier.
    /// </summary>
    [XmlType(Namespace = CT_Presentation.Namespace)]
    public class CT_SlideIdListEntry
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

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

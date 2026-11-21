using System;
using System.Linq;
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
        private int[] elementOrder = Array.Empty<int>();

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlIgnore]
        public XmlElement cSld { get; set; }

        [XmlIgnore]
        public XmlElement clrMap { get; set; }

        [XmlIgnore]
        public XmlElement sldLayoutIdLst { get; set; }

        [XmlIgnore]
        public XmlElement[] AnyElements { get; set; }

        /// <summary>
        /// Gets or sets modeled and opaque master children in their original XML order.
        /// </summary>
        [XmlAnyElement]
        public XmlElement[] Content
        {
            get => PresentationXmlContent.Merge(elementOrder,
                new object[] { cSld }, new object[] { clrMap }, new object[] { sldLayoutIdLst }, AnyElements)
                .Cast<XmlElement>().ToArray();
            set
            {
                var content = value ?? Array.Empty<XmlElement>();
                elementOrder = content.Select(GetElementGroup).ToArray();
                cSld = content.FirstOrDefault(element => GetElementGroup(element) == 0);
                clrMap = content.FirstOrDefault(element => GetElementGroup(element) == 1);
                sldLayoutIdLst = content.FirstOrDefault(element => GetElementGroup(element) == 2);
                AnyElements = content.Where(element => GetElementGroup(element) == 3).ToArray();
            }
        }

        private static int GetElementGroup(XmlElement element)
        {
            if (element.NamespaceURI == CT_Presentation.Namespace)
            {
                switch (element.LocalName)
                {
                    case "cSld": return 0;
                    case "clrMap": return 1;
                    case "sldLayoutIdLst": return 2;
                }
            }
            return 3;
        }
    }
}

using System;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Presentation
{
    /// <summary>
    /// A slide whose drawing content is preserved until high-level shape APIs are available.
    /// </summary>
    [XmlRoot("sld", Namespace = CT_Presentation.Namespace)]
    public class CT_Slide
    {
        private int[] elementOrder = Array.Empty<int>();

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces Namespaces { get; set; }

        [XmlAnyAttribute]
        public XmlAttribute[] AnyAttributes { get; set; }

        [XmlIgnore]
        public XmlElement cSld { get; set; }

        [XmlIgnore]
        public XmlElement[] AnyElements { get; set; }

        /// <summary>
        /// Gets or sets common slide data and opaque children in their original XML order.
        /// </summary>
        [XmlAnyElement]
        public XmlElement[] Content
        {
            get => PresentationXmlContent.Merge(elementOrder, new object[] { cSld }, AnyElements)
                .Cast<XmlElement>().ToArray();
            set
            {
                var content = value ?? Array.Empty<XmlElement>();
                elementOrder = content.Select(element => IsCommonSlideData(element) ? 0 : 1).ToArray();
                cSld = content.FirstOrDefault(IsCommonSlideData);
                AnyElements = content.Where(element => !IsCommonSlideData(element)).ToArray();
            }
        }

        private static bool IsCommonSlideData(XmlElement element)
        {
            return element.LocalName == "cSld" && element.NamespaceURI == CT_Presentation.Namespace;
        }
    }
}

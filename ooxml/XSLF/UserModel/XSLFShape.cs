using System;
using System.Xml;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A shape backed by preserved PresentationML XML.
    /// </summary>
    public abstract class XSLFShape
    {
        private readonly XmlElement xml;
        private readonly XSLFSheet sheet;

        internal XSLFShape(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
        {
            this.xml = xml;
            this.sheet = sheet;
            Parent = parent;
        }

        /// <summary>
        /// Gets the immediate containing sheet or group.
        /// </summary>
        public IXSLFShapeContainer Parent { get; }

        /// <summary>
        /// Gets the stable identifier from the nonvisual drawing properties.
        /// Opaque content without unambiguous nonvisual properties is not interpreted.
        /// </summary>
        public uint ID
        {
            get { return XSLFShapeHelper.ReadId(GetNonVisualProperties()); }
        }

        /// <summary>
        /// Gets or sets the shape name without changing its identifier.
        /// Opaque wrappers do not support mutation.
        /// </summary>
        public string Name
        {
            get
            {
                var properties = GetNonVisualProperties();
                if (!properties.HasAttribute("name"))
                {
                    throw new POIXMLException("Shape nonvisual properties have no name.");
                }
                return properties.GetAttribute("name");
            }
            set
            {
                EnsureEditable();
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }
                XmlConvert.VerifyXmlChars(value);
                GetNonVisualProperties().SetAttribute("name", value);
            }
        }

        /// <summary>
        /// Returns the sheet containing this shape.
        /// </summary>
        public XSLFSheet GetSheet()
        {
            return sheet;
        }

        /// <summary>
        /// Returns the live underlying XML element, including unsupported content.
        /// </summary>
        public XmlElement GetXmlObject()
        {
            return xml;
        }

        /// <summary>
        /// Returns the explicit unrotated anchor in points, relative to the parent coordinate system.
        /// Missing or inherited transforms and opaque drawing types are not resolved.
        /// </summary>
        public virtual SKRect GetAnchor()
        {
            throw new NotSupportedException("This drawing type has no supported anchor.");
        }

        /// <summary>
        /// Sets the explicit anchor in points without scaling children or changing rotation/flips.
        /// Width and height must be nonnegative; positions may be negative.
        /// Each extent must fit 2147483647 EMUs; positions use DrawingML's wider coordinate range.
        /// </summary>
        public virtual void SetAnchor(SKRect anchor)
        {
            throw new NotSupportedException("This drawing type does not support anchor editing.");
        }

        internal virtual void EnsureEditable()
        {
            XSLFShapeHelper.EnsureAttached(sheet, xml);
        }

        private XmlElement GetNonVisualProperties()
        {
            XmlElement properties = null;
            foreach (XmlNode node in xml.ChildNodes)
            {
                if (node is XmlElement child)
                {
                    var candidate = XSLFShapeHelper.Child(child, "cNvPr", XSLFShapeHelper.PresentationNamespace);
                    if (candidate != null)
                    {
                        if (properties != null)
                        {
                            throw new POIXMLException("Shape nonvisual properties are ambiguous.");
                        }
                        properties = candidate;
                    }
                }
            }
            return properties ?? throw new NotSupportedException(
                "This drawing has no directly supported nonvisual properties.");
        }
    }
}

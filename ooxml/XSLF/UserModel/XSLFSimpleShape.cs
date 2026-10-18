using System;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A PresentationML simple shape. Existing geometry, text, styles and extensions are retained.
    /// </summary>
    public class XSLFSimpleShape : XSLFShape
    {
        internal XSLFSimpleShape(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }

        /// <summary>
        /// Gets the preset geometry, or null when the shape has no preset geometry.
        /// Custom geometry and text formatting are not interpreted.
        /// </summary>
        public ST_ShapeType? ShapeType
        {
            get
            {
                var geometry = XSLFShapeHelper.Child(Properties, "prstGeom", XSLFXmlHelper.DrawingNamespace);
                if (geometry == null)
                {
                    return null;
                }
                if (!Enum.TryParse(geometry.GetAttribute("prst"), out ST_ShapeType type) ||
                    !XSLFShapeHelper.IsPreset(type) || type.ToString() != geometry.GetAttribute("prst"))
                {
                    throw new POIXMLException("Invalid preset shape geometry.");
                }
                return type;
            }
        }

        /// <inheritdoc/>
        public override SKRect GetAnchor()
        {
            return XSLFShapeHelper.ReadAnchor(Properties, false);
        }

        /// <inheritdoc/>
        public override void SetAnchor(SKRect anchor)
        {
            EnsureEditable();
            XSLFShapeHelper.SetAnchor(Properties, anchor, false);
        }

        private XmlElement Properties
        {
            get { return XSLFShapeHelper.RequiredChild(GetXmlObject(), "spPr", XSLFShapeHelper.PresentationNamespace); }
        }
    }
}

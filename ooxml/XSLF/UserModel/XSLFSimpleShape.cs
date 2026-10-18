using System;
using System.Collections.Generic;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A PresentationML simple shape with text support and preserved unsupported content.
    /// </summary>
    public class XSLFSimpleShape : XSLFShape, IXSLFTextContainer
    {
        internal XSLFSimpleShape(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }

        /// <summary>
        /// Gets actual text, joining paragraphs and explicit line breaks with LF.
        /// Placeholder prompts are not inherited. Unsupported text content fails explicitly.
        /// </summary>
        public string Text
        {
            get { return XSLFTextHelper.GetText(this); }
        }

        /// <summary>
        /// Gets a read-only snapshot of this shape's explicit paragraphs.
        /// </summary>
        public IReadOnlyList<XSLFTextParagraph> TextParagraphs
        {
            get { return XSLFTextHelper.GetParagraphs(this); }
        }

        /// <summary>
        /// Replaces all paragraphs with unformatted text, splitting CR/LF into paragraphs.
        /// Body properties/list styles are retained. Unsupported content/formatting blocks replacement.
        /// Edit individual runs to retain paragraph/run formatting.
        /// </summary>
        public void SetText(string text)
        {
            XSLFTextHelper.SetText(this, text);
        }

        /// <summary>
        /// Appends an empty paragraph, creating a text body when absent.
        /// </summary>
        public XSLFTextParagraph AddNewTextParagraph()
        {
            return XSLFTextHelper.AddParagraph(this);
        }

        XmlElement IXSLFTextContainer.GetTextBody(bool create)
        {
            return XSLFTextHelper.GetBody(GetXmlObject(), XSLFShapeHelper.PresentationNamespace, create);
        }

        void IXSLFTextContainer.EnsureTextEditable()
        {
            EnsureEditable();
        }

        /// <summary>
        /// Gets the preset geometry, or null when the shape has no preset geometry.
        /// Custom geometry is not interpreted.
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

using System.Collections.Generic;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A group whose child coordinates and exterior anchor are independent.
    /// </summary>
    public class XSLFGroupShape : XSLFShape, IXSLFShapeContainer
    {
        internal XSLFGroupShape(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }

        /// <inheritdoc/>
        public IReadOnlyList<XSLFShape> GetShapes()
        {
            return XSLFShapeHelper.GetShapes(GetSheet(), this, GetXmlObject());
        }

        /// <inheritdoc/>
        public XSLFSimpleShape CreateAutoShape(ST_ShapeType shapeType = ST_ShapeType.rect)
        {
            return XSLFShapeHelper.CreateAutoShape(GetSheet(), this, GetXmlObject(), shapeType);
        }

        /// <inheritdoc/>
        public XSLFGroupShape CreateGroup()
        {
            return XSLFShapeHelper.CreateGroup(GetSheet(), this, GetXmlObject());
        }

        /// <inheritdoc/>
        public XSLFTextBox CreateTextBox()
        {
            return XSLFShapeHelper.CreateTextBox(GetSheet(), this, GetXmlObject());
        }

        /// <inheritdoc/>
        public XSLFPictureShape CreatePicture(XSLFPictureData pictureData)
        {
            return XSLFPictureShape.Create(GetSheet(), this, GetXmlObject(), pictureData);
        }

        /// <inheritdoc/>
        public XSLFTable CreateTable(int rows, int columns)
        {
            return XSLFTable.Create(GetSheet(), this, GetXmlObject(), rows, columns);
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

        /// <summary>
        /// Returns the group's explicit child-coordinate rectangle in points.
        /// Child anchors use this coordinate system, not absolute slide coordinates.
        /// </summary>
        public SKRect GetInteriorAnchor()
        {
            return XSLFShapeHelper.ReadAnchor(Properties, true);
        }

        /// <summary>
        /// Sets the child-coordinate rectangle without modifying any child anchor or the exterior anchor.
        /// Rotation, flips, and unknown transform content are preserved.
        /// </summary>
        public void SetInteriorAnchor(SKRect anchor)
        {
            EnsureEditable();
            XSLFShapeHelper.SetAnchor(Properties, anchor, true);
        }

        private XmlElement Properties
        {
            get { return XSLFShapeHelper.RequiredChild(GetXmlObject(), "grpSpPr", XSLFShapeHelper.PresentationNamespace); }
        }
    }
}

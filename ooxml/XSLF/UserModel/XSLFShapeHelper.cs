using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.Util;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    internal static class XSLFShapeHelper
    {
        internal const string PresentationNamespace = CT_Presentation.Namespace;
        private const long MinCoordinate = -27273042329600L;
        private const long MaxCoordinate = 27273042316900L;

        internal static XmlElement Child(XmlElement parent, string name, string ns)
        {
            if (parent == null)
            {
                throw new POIXMLException("Missing parent for drawing element: " + name);
            }
            XmlElement result = null;
            foreach (XmlNode node in parent.ChildNodes)
            {
                if (node is XmlElement element && element.LocalName == name && element.NamespaceURI == ns)
                {
                    if (result != null)
                    {
                        throw new POIXMLException("Duplicate drawing element: " + name);
                    }
                    result = element;
                }
            }
            return result;
        }

        internal static XmlElement RequiredChild(XmlElement parent, string name, string ns)
        {
            return Child(parent, name, ns) ?? throw new POIXMLException("Missing drawing element: " + name);
        }

        internal static uint ReadId(XmlElement properties)
        {
            if (!uint.TryParse(properties.GetAttribute("id"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var id))
            {
                throw new POIXMLException("Invalid drawing identifier.");
            }
            return id;
        }

        internal static void EnsureAttached(XSLFSheet sheet, XmlElement element)
        {
            sheet.EnsureShapesWritable();
            var tree = sheet.GetShapeTree();
            for (XmlNode node = element; node != null; node = node.ParentNode)
            {
                if (ReferenceEquals(node, tree))
                {
                    return;
                }
            }
            throw new InvalidOperationException("The shape is no longer attached to this sheet.");
        }

        internal static ReadOnlyCollection<XSLFShape> GetShapes(
            XSLFSheet sheet, IXSLFShapeContainer parent, XmlElement tree)
        {
            ValidateContainer(tree);
            var shapes = new List<XSLFShape>();
            foreach (XmlNode node in tree.ChildNodes)
            {
                if (!(node is XmlElement element))
                {
                    continue;
                }
                if (element.NamespaceURI == PresentationNamespace)
                {
                    switch (element.LocalName)
                    {
                        case "nvGrpSpPr":
                        case "grpSpPr":
                        case "extLst":
                            continue;
                        case "sp":
                            shapes.Add(XSLFTextHelper.IsTextBox(element)
                                ? new XSLFTextBox(element, sheet, parent)
                                : new XSLFSimpleShape(element, sheet, parent));
                            continue;
                        case "grpSp":
                            shapes.Add(new XSLFGroupShape(element, sheet, parent));
                            continue;
                        case "pic":
                            var fill = Child(element, "blipFill", PresentationNamespace);
                            if (fill != null && Child(fill, "blip", XSLFXmlHelper.DrawingNamespace) != null &&
                                Child(element, "spPr", PresentationNamespace) != null)
                            {
                                shapes.Add(new XSLFPictureShape(element, sheet, parent));
                                continue;
                            }
                            break;
                        case "graphicFrame":
                            if (XSLFTable.IsTable(element))
                            {
                                shapes.Add(new XSLFTable(element, sheet, parent));
                                continue;
                            }
                            if (XSLFChartShape.IsChart(element))
                            {
                                shapes.Add(new XSLFChartShape(element, sheet, parent));
                                continue;
                            }
                            break;
                    }
                }
                shapes.Add(new XSLFUnsupportedShape(element, sheet, parent));
            }
            return shapes.AsReadOnly();
        }

        internal static bool IsPreset(ST_ShapeType type)
        {
            return Enum.IsDefined(typeof(ST_ShapeType), type) &&
                type != ST_ShapeType.none && type != ST_ShapeType.lineInv;
        }

        internal static XSLFSimpleShape CreateAutoShape(
            XSLFSheet sheet, IXSLFShapeContainer parent, XmlElement tree, ST_ShapeType type)
        {
            EnsureAttached(sheet, tree);
            if (!IsPreset(type))
            {
                throw new ArgumentOutOfRangeException(nameof(type), "A valid DrawingML preset shape is required.");
            }
            var id = NextId(sheet);
            var xml = XSLFXmlHelper.CreateElement(
                "<p:sp xmlns:p=\"" + PresentationNamespace + "\" xmlns:a=\"" + XSLFXmlHelper.DrawingNamespace +
                "\"><p:nvSpPr><p:cNvPr id=\"" + id.ToString(CultureInfo.InvariantCulture) +
                "\" name=\"AutoShape " + id.ToString(CultureInfo.InvariantCulture) +
                "\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>" +
                "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>" +
                "<a:prstGeom prst=\"" + type + "\"><a:avLst/></a:prstGeom></p:spPr></p:sp>");
            return new XSLFSimpleShape(Append(tree, xml), sheet, parent);
        }

        internal static XSLFGroupShape CreateGroup(XSLFSheet sheet, IXSLFShapeContainer parent, XmlElement tree)
        {
            EnsureAttached(sheet, tree);
            var id = NextId(sheet);
            var xml = XSLFXmlHelper.CreateElement(
                "<p:grpSp xmlns:p=\"" + PresentationNamespace + "\" xmlns:a=\"" + XSLFXmlHelper.DrawingNamespace +
                "\"><p:nvGrpSpPr><p:cNvPr id=\"" + id.ToString(CultureInfo.InvariantCulture) +
                "\" name=\"Group " + id.ToString(CultureInfo.InvariantCulture) +
                "\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>" +
                "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>" +
                "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr></p:grpSp>");
            return new XSLFGroupShape(Append(tree, xml), sheet, parent);
        }

        internal static XSLFTextBox CreateTextBox(XSLFSheet sheet, IXSLFShapeContainer parent, XmlElement tree)
        {
            var shape = CreateAutoShape(sheet, parent, tree, ST_ShapeType.rect);
            var xml = shape.GetXmlObject();
            RequiredChild(RequiredChild(xml, "nvSpPr", PresentationNamespace), "cNvSpPr", PresentationNamespace)
                .SetAttribute("txBox", "1");
            shape.Name = "TextBox " + shape.ID.ToString(CultureInfo.InvariantCulture);
            var properties = RequiredChild(xml, "spPr", PresentationNamespace);
            properties.AppendChild(xml.OwnerDocument.CreateElement("a", "noFill", XSLFXmlHelper.DrawingNamespace));
            var line = xml.OwnerDocument.CreateElement("a", "ln", XSLFXmlHelper.DrawingNamespace);
            line.AppendChild(xml.OwnerDocument.CreateElement("a", "noFill", XSLFXmlHelper.DrawingNamespace));
            properties.AppendChild(line);
            var textBox = new XSLFTextBox(xml, sheet, parent);
            textBox.AddNewTextParagraph();
            return textBox;
        }

        internal static SKRect ReadAnchor(XmlElement properties, bool interior,
            string transformNamespace = XSLFXmlHelper.DrawingNamespace)
        {
            var transform = Child(properties, "xfrm", transformNamespace);
            var off = transform == null ? null : Child(transform, interior ? "chOff" : "off", XSLFXmlHelper.DrawingNamespace);
            var ext = transform == null ? null : Child(transform, interior ? "chExt" : "ext", XSLFXmlHelper.DrawingNamespace);
            if (off == null || ext == null)
            {
                throw new NotSupportedException("No explicit anchor is available; inherited transforms are not resolved.");
            }
            var x = ReadCoordinate(off, "x", false);
            var y = ReadCoordinate(off, "y", false);
            var width = ReadCoordinate(ext, "cx", true);
            var height = ReadCoordinate(ext, "cy", true);
            return SKRect.Create((float)Units.ToPoints(x), (float)Units.ToPoints(y),
                (float)Units.ToPoints(width), (float)Units.ToPoints(height));
        }

        internal static void SetAnchor(XmlElement properties, SKRect anchor, bool interior,
            string transformNamespace = XSLFXmlHelper.DrawingNamespace)
        {
            var x = ToCoordinate(anchor.Left, false);
            var y = ToCoordinate(anchor.Top, false);
            var width = ToCoordinate(anchor.Width, true);
            var height = ToCoordinate(anchor.Height, true);
            var transform = Child(properties, "xfrm", transformNamespace);
            var offName = interior ? "chOff" : "off";
            var extName = interior ? "chExt" : "ext";
            var off = transform == null ? null : Child(transform, offName, XSLFXmlHelper.DrawingNamespace);
            var ext = transform == null ? null : Child(transform, extName, XSLFXmlHelper.DrawingNamespace);
            if (transform == null)
            {
                if (transformNamespace != XSLFXmlHelper.DrawingNamespace)
                    throw new POIXMLException("Missing graphic frame transform.");
                transform = properties.OwnerDocument.CreateElement("a", "xfrm", XSLFXmlHelper.DrawingNamespace);
                properties.PrependChild(transform);
            }
            if (off == null)
            {
                off = NewTransformChild(transform, offName);
            }
            if (ext == null)
            {
                ext = NewTransformChild(transform, extName);
            }
            off.SetAttribute("x", x.ToString(CultureInfo.InvariantCulture));
            off.SetAttribute("y", y.ToString(CultureInfo.InvariantCulture));
            ext.SetAttribute("cx", width.ToString(CultureInfo.InvariantCulture));
            ext.SetAttribute("cy", height.ToString(CultureInfo.InvariantCulture));
        }

        private static XmlElement NewTransformChild(XmlElement transform, string name)
        {
            var order = new[] { "off", "ext", "chOff", "chExt" };
            var index = Array.IndexOf(order, name);
            var following = transform.ChildNodes.OfType<XmlElement>().FirstOrDefault(child =>
                child.NamespaceURI == XSLFXmlHelper.DrawingNamespace &&
                Array.IndexOf(order, child.LocalName) > index);
            var element = transform.OwnerDocument.CreateElement("a", name, XSLFXmlHelper.DrawingNamespace);
            transform.InsertBefore(element, following);
            return element;
        }

        private static long ReadCoordinate(XmlElement element, string attribute, bool extent)
        {
            if (!long.TryParse(element.GetAttribute(attribute), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var value) ||
                value < (extent ? 0 : MinCoordinate) || value > (extent ? int.MaxValue : MaxCoordinate))
            {
                throw new POIXMLException("Invalid or unsupported DrawingML coordinate: " + attribute);
            }
            return value;
        }

        private static long ToCoordinate(float points, bool extent)
        {
            var value = Math.Round((double)points * Units.EMU_PER_POINT);
            if (float.IsNaN(points) || float.IsInfinity(points) || (extent && points < 0) ||
                value < (extent ? 0 : MinCoordinate) || value > (extent ? int.MaxValue : MaxCoordinate))
            {
                throw new ArgumentOutOfRangeException(nameof(points),
                    "Shape coordinates must be finite DrawingML coordinates with nonnegative extents.");
            }
            return (long)value;
        }

        internal static uint NextId(XSLFSheet sheet)
        {
            var ids = new HashSet<uint>();
            foreach (XmlElement properties in sheet.GetShapeTree().GetElementsByTagName("cNvPr", PresentationNamespace))
            {
                ids.Add(ReadId(properties));
            }
            var id = 1U;
            while (ids.Contains(id))
            {
                if (id == uint.MaxValue)
                {
                    throw new InvalidOperationException("No drawing identifier is available.");
                }
                id++;
            }
            return id;
        }

        internal static XmlElement Append(XmlElement tree, XmlElement xml)
        {
            ValidateContainer(tree);
            var extension = Child(tree, "extLst", PresentationNamespace);
            var element = (XmlElement)tree.OwnerDocument.ImportNode(xml, true);
            tree.InsertBefore(element, extension);
            return element;
        }

        internal static void ValidateContainer(XmlElement tree)
        {
            var nonvisual = RequiredChild(tree, "nvGrpSpPr", PresentationNamespace);
            ReadId(RequiredChild(nonvisual, "cNvPr", PresentationNamespace));
            RequiredChild(tree, "grpSpPr", PresentationNamespace);
            Child(tree, "extLst", PresentationNamespace);
        }
    }
}

using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml;

namespace NPOI.XSLF.UserModel
{
    internal static class XSLFTextHelper
    {
        internal const string A = XSLFXmlHelper.DrawingNamespace;
        internal const string P = XSLFShapeHelper.PresentationNamespace;
        private const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";

        internal static bool IsTextBox(XmlElement shape)
        {
            var nv = XSLFShapeHelper.Child(shape, "nvSpPr", P);
            var properties = nv == null ? null : XSLFShapeHelper.Child(nv, "cNvSpPr", P);
            return properties != null && properties.HasAttribute("txBox") &&
                ParseBoolean(properties.GetAttribute("txBox"));
        }

        internal static bool ParseBoolean(string value)
        {
            switch (value.Trim())
            {
                case "true":
                case "1":
                    return true;
                case "false":
                case "0":
                    return false;
                default:
                    throw new POIXMLException("Invalid DrawingML boolean: " + value);
            }
        }

        internal static T ParseEnum<T>(string value) where T : struct, Enum
        {
            if (!Enum.TryParse(value, out T result) || !Enum.IsDefined(typeof(T), result) || result.ToString() != value)
                throw new POIXMLException("Invalid DrawingML enumeration: " + value);
            return result;
        }

        internal static void VerifyText(string text)
        {
            if (text == null)
                throw new ArgumentNullException(nameof(text));
            XmlConvert.VerifyXmlChars(text);
        }

        internal static XmlElement GetBody(IXSLFTextContainer owner, bool create)
        {
            return owner.GetTextBody(create);
        }

        internal static XmlElement GetBody(XmlElement xml, string ns, bool create)
        {
            var body = XSLFShapeHelper.Child(xml, "txBody", ns);
            if (body != null)
            {
                XSLFShapeHelper.RequiredChild(body, "bodyPr", A);
                XSLFShapeHelper.Child(body, "lstStyle", A);
            }
            else if (create)
            {
                var following = XSLFShapeHelper.Child(xml, "extLst", ns);
                if (ns == A)
                    following = XSLFShapeHelper.Child(xml, "tcPr", A) ?? following;
                body = xml.OwnerDocument.CreateElement(ns == P ? "p" : "a", "txBody", ns);
                body.AppendChild(xml.OwnerDocument.CreateElement("a", "bodyPr", A));
                body.AppendChild(xml.OwnerDocument.CreateElement("a", "lstStyle", A));
                xml.InsertBefore(body, following);
            }
            return body;
        }

        internal static ReadOnlyCollection<XSLFTextParagraph> GetParagraphs(IXSLFTextContainer shape)
        {
            var body = GetBody(shape, false);
            return (body == null ? Enumerable.Empty<XmlElement>() :
                body.ChildNodes.OfType<XmlElement>().Where(node => node.NamespaceURI == A && node.LocalName == "p"))
                .Select(node => new XSLFTextParagraph(node, shape)).ToList().AsReadOnly();
        }

        internal static ReadOnlyCollection<XSLFTextRun> GetRuns(XSLFTextParagraph paragraph)
        {
            XSLFShapeHelper.Child(paragraph.GetXmlObject(), "pPr", A);
            XSLFShapeHelper.Child(paragraph.GetXmlObject(), "endParaRPr", A);
            return paragraph.GetXmlObject().ChildNodes.OfType<XmlElement>().Where(node =>
                node.NamespaceURI != A || (node.LocalName != "pPr" && node.LocalName != "endParaRPr"))
                .Select(node => new XSLFTextRun(node, paragraph)).ToList().AsReadOnly();
        }

        internal static string GetText(IXSLFTextContainer shape)
        {
            var body = GetBody(shape, false);
            if (body == null)
                return string.Empty;
            CheckBodyChildren(body);
            return string.Join("\n", GetParagraphs(shape).Select(paragraph => paragraph.Text));
        }

        internal static XSLFTextParagraph AddParagraph(IXSLFTextContainer shape)
        {
            shape.EnsureTextEditable();
            var body = GetBody(shape, false);
            if (body != null)
                CheckBodyChildren(body);
            body = body ?? GetBody(shape, true);
            var paragraph = body.OwnerDocument.CreateElement("a", "p", A);
            body.AppendChild(paragraph);
            return new XSLFTextParagraph(paragraph, shape);
        }

        internal static XSLFTextRun AddRun(XSLFTextParagraph paragraph, bool lineBreak)
        {
            paragraph.EnsureEditable();
            var xml = paragraph.GetXmlObject();
            var end = XSLFShapeHelper.Child(xml, "endParaRPr", A);
            var run = xml.OwnerDocument.CreateElement("a", lineBreak ? "br" : "r", A);
            if (!lineBreak)
            {
                var text = xml.OwnerDocument.CreateElement("a", "t", A);
                SetTextElement(text, string.Empty);
                run.AppendChild(text);
            }
            xml.InsertBefore(run, end);
            return new XSLFTextRun(run, paragraph);
        }

        internal static void SetText(IXSLFTextContainer shape, string text)
        {
            shape.EnsureTextEditable();
            VerifyText(text);
            var body = GetBody(shape, false);
            if (body != null)
            {
                CheckBodyChildren(body);
                foreach (var paragraph in GetParagraphs(shape))
                {
                    if (!CanReplace(paragraph.GetXmlObject()))
                        throw new NotSupportedException(
                            "Whole-text replacement would discard unsupported content or formatting; edit regular runs instead.");
                }
            }
            var document = shape.GetXmlObject().OwnerDocument;
            var paragraphs = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Select(value =>
            {
                var paragraph = document.CreateElement("a", "p", A);
                var run = document.CreateElement("a", "r", A);
                var element = document.CreateElement("a", "t", A);
                SetTextElement(element, value);
                run.AppendChild(element);
                paragraph.AppendChild(run);
                return paragraph;
            }).ToArray();
            body = body ?? GetBody(shape, true);
            foreach (var paragraph in GetParagraphs(shape))
                body.RemoveChild(paragraph.GetXmlObject());
            foreach (var paragraph in paragraphs)
                body.AppendChild(paragraph);
        }

        internal static void SetTextElement(XmlElement element, string text)
        {
            var space = element.OwnerDocument.CreateAttribute("xml", "space", XmlNamespace);
            space.Value = "preserve";
            element.SetAttributeNode(space);
            element.InnerText = text;
        }

        internal static XmlElement ParagraphProperties(XSLFTextParagraph paragraph, bool create)
        {
            var xml = paragraph.GetXmlObject();
            var properties = XSLFShapeHelper.Child(xml, "pPr", A);
            if (properties == null && create)
            {
                properties = xml.OwnerDocument.CreateElement("a", "pPr", A);
                xml.PrependChild(properties);
            }
            return properties;
        }

        internal static XmlElement RunProperties(XSLFTextRun run, bool create)
        {
            if (run.Kind != "r" && run.Kind != "br")
                throw new NotSupportedException("Field and opaque inline formatting is not interpreted.");
            CheckRunChildren(run);
            var xml = run.GetXmlObject();
            var properties = XSLFShapeHelper.Child(xml, "rPr", A);
            if (properties == null && create)
            {
                properties = xml.OwnerDocument.CreateElement("a", "rPr", A);
                xml.PrependChild(properties);
            }
            return properties;
        }

        internal static void CheckRunChildren(XSLFTextRun run)
        {
            foreach (XmlElement child in run.GetXmlObject().ChildNodes.OfType<XmlElement>())
            {
                if (child.NamespaceURI != A || (child.LocalName != "rPr" &&
                    !(child.LocalName == "t" && (run.Kind == "r" || run.Kind == "fld")) &&
                    !(child.LocalName == "pPr" && run.Kind == "fld")))
                    throw new NotSupportedException("Opaque content inside a text run is not interpreted or edited.");
            }
        }

        internal static XmlElement Insert(XmlElement parent, string name, string[] followingNames)
        {
            var following = parent.ChildNodes.OfType<XmlElement>().FirstOrDefault(node =>
                node.NamespaceURI == A && followingNames.Contains(node.LocalName));
            var element = parent.OwnerDocument.CreateElement("a", name, A);
            parent.InsertBefore(element, following);
            return element;
        }

        private static void CheckBodyChildren(XmlElement body)
        {
            if (body.ChildNodes.OfType<XmlElement>().Any(node =>
                node.NamespaceURI != A || (node.LocalName != "bodyPr" && node.LocalName != "lstStyle" && node.LocalName != "p")))
                throw new NotSupportedException("Opaque text-body content is not interpreted or replaced.");
        }

        private static bool CanReplace(XmlElement element)
        {
            if (element.NamespaceURI != A)
                return false;
            string[] attributes;
            string[] children;
            switch (element.LocalName)
            {
                case "p":
                    attributes = Array.Empty<string>();
                    children = new[] { "pPr", "r", "br", "endParaRPr" };
                    XSLFShapeHelper.Child(element, "pPr", A);
                    XSLFShapeHelper.Child(element, "endParaRPr", A);
                    break;
                case "r":
                    attributes = Array.Empty<string>();
                    children = new[] { "rPr", "t" };
                    XSLFShapeHelper.RequiredChild(element, "t", A);
                    XSLFShapeHelper.Child(element, "rPr", A);
                    break;
                case "br":
                    attributes = Array.Empty<string>();
                    children = new[] { "rPr" };
                    XSLFShapeHelper.Child(element, "rPr", A);
                    break;
                case "t":
                    attributes = Array.Empty<string>();
                    children = Array.Empty<string>();
                    break;
                case "pPr":
                    attributes = new[] { "lvl", "algn" };
                    children = new[] { "defRPr" };
                    XSLFShapeHelper.Child(element, "defRPr", A);
                    break;
                case "rPr":
                case "defRPr":
                case "endParaRPr":
                    attributes = new[] { "sz", "b", "i", "u" };
                    children = new[] { "latin" };
                    XSLFShapeHelper.Child(element, "latin", A);
                    break;
                case "latin":
                    attributes = new[] { "typeface" };
                    children = Array.Empty<string>();
                    break;
                default:
                    return false;
            }
            foreach (XmlAttribute attribute in element.Attributes)
            {
                if (attribute.NamespaceURI == "http://www.w3.org/2000/xmlns/" ||
                    (element.LocalName == "t" && attribute.NamespaceURI == XmlNamespace && attribute.LocalName == "space"))
                    continue;
                if (attribute.NamespaceURI.Length != 0 || !attributes.Contains(attribute.LocalName))
                    return false;
            }
            foreach (XmlNode child in element.ChildNodes)
            {
                if (child is XmlElement node)
                {
                    if (!children.Contains(node.LocalName) || !CanReplace(node))
                        return false;
                }
                else if (!(child is XmlWhitespace) &&
                    !(element.LocalName == "t" && (child is XmlText || child is XmlCDataSection || child is XmlSignificantWhitespace)))
                    return false;
            }
            return true;
        }
    }
}

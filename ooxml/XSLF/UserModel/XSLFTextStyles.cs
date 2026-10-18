using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace NPOI.XSLF.UserModel
{
    internal static class XSLFTextStyles
    {
        private const string A = XSLFTextHelper.A;
        private const string P = XSLFTextHelper.P;

        internal static IEnumerable<XmlElement> RunProperties(XSLFTextRun run)
        {
            var local = XSLFTextHelper.RunProperties(run, false);
            if (local != null)
                yield return local;
            foreach (var properties in ParagraphProperties(run.ParentParagraph))
            {
                var defaults = XSLFShapeHelper.Child(properties, "defRPr", A);
                if (defaults != null)
                    yield return defaults;
            }
        }

        internal static IEnumerable<XmlElement> ParagraphProperties(XSLFTextParagraph paragraph)
        {
            var local = XSLFTextHelper.ParagraphProperties(paragraph, false);
            if (local != null)
                yield return local;
            var level = paragraph.IndentLevel;
            foreach (var properties in ShapeListProperties(paragraph.Owner, level))
                yield return properties;
            if (paragraph.ParentCell != null)
                yield break;
            var shape = paragraph.ParentShape;
            var sheet = shape.GetSheet();
            var placeholder = Placeholder(shape);
            if (placeholder != null)
            {
                if (sheet is XSLFSlide slide)
                {
                    var layout = FindPlaceholder(slide.GetSlideLayout(), placeholder, true);
                    if (layout != null)
                    {
                        foreach (var properties in ShapeListProperties(layout, level))
                            yield return properties;
                    }
                }
                if (!(sheet is XSLFSlideMaster))
                {
                    var masterShape = FindPlaceholder(Master(sheet), placeholder, false);
                    if (masterShape != null)
                    {
                        foreach (var properties in ShapeListProperties(masterShape, level))
                            yield return properties;
                    }
                }
            }
            else
            {
                var defaults = Root(sheet.GetSlideShow().GetCTPresentation().AnyElements, "defaultTextStyle");
                foreach (var properties in ListProperties(defaults, level))
                    yield return properties;
            }
            var master = Master(sheet);
            var styles = Root(master.GetXmlObject().AnyElements, "txStyles");
            var type = placeholder == null ? null : PlaceholderType(placeholder);
            var selector = type == "title" || type == "ctrTitle" ? "titleStyle" :
                type == null || type == "dt" || type == "ftr" || type == "hdr" || type == "sldNum" ? "otherStyle" : "bodyStyle";
            var style = styles == null ? null : XSLFShapeHelper.Child(styles, selector, P);
            foreach (var properties in ListProperties(style, level))
                yield return properties;
        }

        internal static string ResolveFont(string font, IXSLFTextContainer shape)
        {
            if (font == "+mj-lt" || font == "+mn-lt")
            {
                var theme = shape.GetSheet().GetTheme();
                var resolved = font == "+mj-lt" ? theme?.MajorFont : theme?.MinorFont;
                if (string.IsNullOrEmpty(resolved))
                    throw new POIXMLException("Cannot resolve the Latin theme font: " + font);
                return resolved;
            }
            if (string.IsNullOrEmpty(font) || font[0] == '+')
                throw new NotSupportedException("Only explicit Latin font families and +mj-lt/+mn-lt are resolved.");
            return font;
        }

        private static IEnumerable<XmlElement> ShapeListProperties(IXSLFTextContainer shape, int level)
        {
            var body = XSLFTextHelper.GetBody(shape, false);
            var style = body == null ? null : XSLFShapeHelper.Child(body, "lstStyle", A);
            return ListProperties(style, level);
        }

        private static IEnumerable<XmlElement> ListProperties(XmlElement style, int level)
        {
            if (style == null)
                yield break;
            var properties = XSLFShapeHelper.Child(style, "lvl" + (level + 1).ToString(CultureInfo.InvariantCulture) + "pPr", A);
            if (properties != null)
                yield return properties;
            properties = XSLFShapeHelper.Child(style, "defPPr", A);
            if (properties != null)
                yield return properties;
        }

        private static XSLFSlideMaster Master(XSLFSheet sheet)
        {
            if (sheet is XSLFSlideMaster master)
                return master;
            if (sheet is XSLFSlideLayout layout)
                return layout.GetSlideMaster();
            if (sheet is XSLFSlide slide)
                return slide.GetSlideMaster();
            throw new NotSupportedException("Text defaults are not supported for this sheet type.");
        }

        private static XmlElement Placeholder(XSLFSimpleShape shape)
        {
            var nv = XSLFShapeHelper.Child(shape.GetXmlObject(), "nvSpPr", P);
            var properties = nv == null ? null : XSLFShapeHelper.Child(nv, "nvPr", P);
            return properties == null ? null : XSLFShapeHelper.Child(properties, "ph", P);
        }

        private static XSLFSimpleShape FindPlaceholder(XSLFSheet sheet, XmlElement source, bool byIndex)
        {
            var sourceType = MasterType(PlaceholderType(source));
            var sourceIndex = byIndex ? Index(source) : 0;
            XSLFSimpleShape result = null;
            foreach (var shape in Shapes(sheet))
            {
                var placeholder = Placeholder(shape);
                if (placeholder == null)
                    continue;
                var matches = byIndex ? Index(placeholder) == sourceIndex :
                    MasterType(PlaceholderType(placeholder)) == sourceType;
                if (!matches)
                    continue;
                if (result != null)
                    throw new POIXMLException("Text placeholder matching is ambiguous.");
                result = shape;
            }
            return result;
        }

        private static IEnumerable<XSLFSimpleShape> Shapes(IXSLFShapeContainer container)
        {
            foreach (var shape in container.GetShapes())
            {
                if (shape is XSLFSimpleShape simple)
                    yield return simple;
                else if (shape is XSLFGroupShape group)
                {
                    foreach (var child in Shapes(group))
                        yield return child;
                }
                else if (shape.GetXmlObject().GetElementsByTagName("ph", P).Count > 0)
                    throw new NotSupportedException("Opaque placeholder content is not resolved.");
            }
        }

        private static uint Index(XmlElement placeholder)
        {
            if (!placeholder.HasAttribute("idx"))
                return 0;
            if (!uint.TryParse(placeholder.GetAttribute("idx"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var index))
                throw new POIXMLException("Invalid placeholder index.");
            return index;
        }

        private static string PlaceholderType(XmlElement placeholder)
        {
            var type = placeholder.HasAttribute("type") ? placeholder.GetAttribute("type") : "obj";
            switch (type)
            {
                case "title": case "ctrTitle": case "body": case "subTitle": case "dt": case "sldNum":
                case "ftr": case "hdr": case "obj": case "chart": case "tbl": case "clipArt": case "dgm":
                case "media": case "pic":
                    return type;
                default:
                    throw new POIXMLException("Invalid placeholder type.");
            }
        }

        private static string MasterType(string type)
        {
            switch (type)
            {
                case "title": case "ctrTitle":
                    return "title";
                case "dt": case "ftr": case "hdr": case "sldNum":
                    return type;
                default:
                    return "body";
            }
        }

        private static XmlElement Root(XmlElement[] elements, string name)
        {
            var matches = (elements ?? Array.Empty<XmlElement>())
                .Where(element => element.NamespaceURI == P && element.LocalName == name).ToArray();
            if (matches.Length > 1)
                throw new POIXMLException("Duplicate presentation text style: " + name);
            return matches.FirstOrDefault();
        }
    }
}

using System;
using System.Globalization;
using System.Linq;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A regular text run, explicit break, displayed field, or opaque inline element.
    /// Field and opaque mutation is not supported.
    /// </summary>
    public class XSLFTextRun
    {
        private static readonly string[] AfterLatin = { "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst" };
        private readonly XmlElement xml;

        internal XSLFTextRun(XmlElement xml, XSLFTextParagraph paragraph)
        {
            this.xml = xml;
            ParentParagraph = paragraph;
        }

        /// <summary>
        /// Gets the containing paragraph.
        /// </summary>
        public XSLFTextParagraph ParentParagraph { get; }

        /// <summary>
        /// Gets displayed text, or LF for a break. Setting text supports regular runs only.
        /// Use AddLineBreak or shape SetText for line/paragraph breaks.
        /// </summary>
        public string Text
        {
            get
            {
                XSLFTextHelper.CheckRunChildren(this);
                if (Kind == "br")
                    return "\n";
                if (Kind != "r" && Kind != "fld")
                    throw new NotSupportedException("Opaque inline text is not interpreted.");
                var text = XSLFShapeHelper.RequiredChild(xml, "t", XSLFTextHelper.A);
                if (text.ChildNodes.OfType<XmlElement>().Any())
                    throw new NotSupportedException("Nested text XML is not interpreted.");
                return text.InnerText;
            }
            set
            {
                EnsureEditable();
                if (Kind != "r")
                    throw new NotSupportedException("Only regular run text can be replaced.");
                XSLFTextHelper.CheckRunChildren(this);
                XSLFTextHelper.VerifyText(value);
                if (value.Contains('\r') || value.Contains('\n'))
                    throw new ArgumentException("Use AddLineBreak or shape SetText for line breaks.", nameof(value));
                var text = XSLFShapeHelper.RequiredChild(xml, "t", XSLFTextHelper.A);
                if (text.ChildNodes.Cast<XmlNode>().Any(node =>
                    !(node is XmlText || node is XmlWhitespace || node is XmlSignificantWhitespace || node is XmlCDataSection)))
                    throw new NotSupportedException("Nested text XML is not supported.");
                XSLFTextHelper.SetTextElement(text, value);
            }
        }

        /// <summary>
        /// Gets or sets the supported effective size in points (1 through 4000).
        /// Values round to hundredths of a point; null removes the local override.
        /// </summary>
        public double? FontSize
        {
            get
            {
                var value = Attribute("sz");
                if (value == null)
                    return null;
                if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var size) ||
                    size < 100 || size > 400000)
                    throw new POIXMLException("Invalid text font size.");
                return size / 100.0;
            }
            set
            {
                EnsureEditable();
                if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value) ||
                    value.Value < 1 || value.Value > 4000))
                    throw new ArgumentOutOfRangeException(nameof(value));
                SetAttribute("sz", value.HasValue
                    ? Math.Round(value.Value * 100).ToString(CultureInfo.InvariantCulture) : null);
            }
        }

        /// <summary>
        /// Gets or sets the supported effective Latin font family.
        /// Resolves +mj-lt/+mn-lt; unsupported or unresolved automatic fonts fail explicitly.
        /// Setting null removes the local Latin override.
        /// </summary>
        public string FontFamily
        {
            get
            {
                foreach (var properties in XSLFTextStyles.RunProperties(this))
                {
                    var latin = XSLFShapeHelper.Child(properties, "latin", XSLFTextHelper.A);
                    if (latin == null)
                        continue;
                    if (!latin.HasAttribute("typeface"))
                        throw new POIXMLException("Latin text font has no typeface.");
                    return XSLFTextStyles.ResolveFont(latin.GetAttribute("typeface"), ParentParagraph.ParentShape);
                }
                return null;
            }
            set
            {
                EnsureEditable();
                if (value != null)
                {
                    XSLFTextHelper.VerifyText(value);
                    if (string.IsNullOrWhiteSpace(value))
                        throw new ArgumentException("A nonempty font family is required.", nameof(value));
                    if (value[0] == '+' && value != "+mj-lt" && value != "+mn-lt")
                        throw new NotSupportedException("Only Latin theme-font tokens are supported.");
                }
                var properties = XSLFTextHelper.RunProperties(this, false);
                var latin = properties == null ? null : XSLFShapeHelper.Child(properties, "latin", XSLFTextHelper.A);
                if (value == null)
                {
                    if (latin != null)
                        properties.RemoveChild(latin);
                    return;
                }
                properties = properties ?? XSLFTextHelper.RunProperties(this, true);
                if (latin == null)
                    latin = XSLFTextHelper.Insert(properties, "latin", AfterLatin);
                latin.SetAttribute("typeface", value);
            }
        }

        /// <summary>
        /// Gets effective bold, or null when unspecified; null removes only the local override.
        /// </summary>
        public bool? IsBold
        {
            get { return BooleanAttribute("b"); }
            set { SetBooleanAttribute("b", value); }
        }

        /// <summary>
        /// Gets effective italic, or null when unspecified; null removes only the local override.
        /// </summary>
        public bool? IsItalic
        {
            get { return BooleanAttribute("i"); }
            set { SetBooleanAttribute("i", value); }
        }

        /// <summary>
        /// Gets whether effective underline is non-none. True sets single underline, false sets none,
        /// and null removes the local override. Other existing underline styles remain untouched.
        /// </summary>
        public bool? IsUnderline
        {
            get
            {
                var value = Attribute("u");
                return value == null ? (bool?)null :
                    XSLFTextHelper.ParseEnum<ST_TextUnderlineType>(value) != ST_TextUnderlineType.none;
            }
            set
            {
                EnsureEditable();
                SetAttribute("u", value.HasValue ? (value.Value ? "sng" : "none") : null);
            }
        }

        /// <summary>
        /// Returns the live underlying inline XML, including field or unknown metadata.
        /// </summary>
        public XmlElement GetXmlObject()
        {
            return xml;
        }

        internal string Kind
        {
            get { return xml.NamespaceURI == XSLFTextHelper.A ? xml.LocalName : string.Empty; }
        }

        internal void EnsureEditable()
        {
            ParentParagraph.EnsureEditable();
            if (!ReferenceEquals(xml.ParentNode, ParentParagraph.GetXmlObject()))
                throw new InvalidOperationException("The text run is detached from its paragraph.");
            if (Kind != "r" && Kind != "br")
                throw new NotSupportedException("Field and opaque inline editing is not supported.");
        }

        private string Attribute(string name)
        {
            foreach (var properties in XSLFTextStyles.RunProperties(this))
            {
                if (properties.HasAttribute(name))
                    return properties.GetAttribute(name);
            }
            return null;
        }

        private bool? BooleanAttribute(string name)
        {
            var value = Attribute(name);
            return value == null ? (bool?)null : XSLFTextHelper.ParseBoolean(value);
        }

        private void SetBooleanAttribute(string name, bool? value)
        {
            EnsureEditable();
            SetAttribute(name, value.HasValue ? (value.Value ? "1" : "0") : null);
        }

        private void SetAttribute(string name, string value)
        {
            var properties = XSLFTextHelper.RunProperties(this, value != null);
            if (properties == null)
                return;
            if (value == null)
                properties.RemoveAttribute(name);
            else
                properties.SetAttribute(name, value);
        }
    }
}

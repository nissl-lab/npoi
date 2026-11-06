using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;
using NPOI.XDDF.UserModel.Text;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A text paragraph with explicit content and individually inherited supported properties.
    /// </summary>
    public class XSLFTextParagraph
    {
        private readonly XmlElement xml;

        internal XSLFTextParagraph(XmlElement xml, XSLFSimpleShape shape)
        {
            this.xml = xml;
            ParentShape = shape;
        }

        /// <summary>
        /// Gets the containing simple shape.
        /// </summary>
        public XSLFSimpleShape ParentShape { get; }

        /// <summary>
        /// Gets actual paragraph text, including displayed field text and explicit LF line breaks.
        /// Fields are not recalculated and opaque text fails explicitly.
        /// </summary>
        public string Text
        {
            get { return string.Concat(TextRuns.Select(run => run.Text)); }
        }

        /// <summary>
        /// Gets a read-only snapshot of runs, fields, breaks and opaque inline elements in XML order.
        /// </summary>
        public IReadOnlyList<XSLFTextRun> TextRuns
        {
            get { return XSLFTextHelper.GetRuns(this); }
        }

        /// <summary>
        /// Gets or sets the zero-based list level, from zero through eight.
        /// An absent level uses the schema's level-zero default.
        /// </summary>
        public int IndentLevel
        {
            get
            {
                var properties = XSLFTextHelper.ParagraphProperties(this, false);
                if (properties == null || !properties.HasAttribute("lvl"))
                    return 0;
                if (!int.TryParse(properties.GetAttribute("lvl"), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out var value) || value < 0 || value > 8)
                    throw new POIXMLException("Invalid text paragraph level.");
                return value;
            }
            set
            {
                EnsureEditable();
                if (value < 0 || value > 8)
                    throw new ArgumentOutOfRangeException(nameof(value));
                XSLFTextHelper.ParagraphProperties(this, true).SetAttribute("lvl",
                    value.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Gets the supported effective alignment, or null when unspecified.
        /// Setting null removes only the local override.
        /// </summary>
        public TextAlignment? TextAlign
        {
            get
            {
                foreach (var properties in XSLFTextStyles.ParagraphProperties(this))
                {
                    if (properties.HasAttribute("algn"))
                    {
                        var value = XSLFTextHelper.ParseEnum<ST_TextAlignType>(properties.GetAttribute("algn"));
                        return TextAlignmentExtensions.ValueOf(value);
                    }
                }
                return null;
            }
            set
            {
                EnsureEditable();
                var alignment = value?.ToST_TextAlignType().ToString();
                var properties = XSLFTextHelper.ParagraphProperties(this, value.HasValue);
                if (properties == null)
                    return;
                if (alignment == null)
                    properties.RemoveAttribute("algn");
                else
                    properties.SetAttribute("algn", alignment);
            }
        }

        /// <summary>
        /// Returns the live underlying paragraph XML.
        /// </summary>
        public XmlElement GetXmlObject()
        {
            return xml;
        }

        /// <summary>
        /// Appends an empty regular run before any end-of-paragraph properties.
        /// </summary>
        public XSLFTextRun AddNewTextRun()
        {
            return XSLFTextHelper.AddRun(this, false);
        }

        /// <summary>
        /// Appends an explicit line break without creating a new paragraph.
        /// </summary>
        public XSLFTextRun AddLineBreak()
        {
            return XSLFTextHelper.AddRun(this, true);
        }

        internal void EnsureEditable()
        {
            ParentShape.EnsureEditable();
            if (!ReferenceEquals(xml.ParentNode, XSLFTextHelper.GetBody(ParentShape, false)))
                throw new InvalidOperationException("The text paragraph is detached from its shape.");
        }
    }
}

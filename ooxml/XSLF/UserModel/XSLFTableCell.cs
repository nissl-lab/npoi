using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A physical table cell using the shared paragraph/run APIs.
    /// Only explicit cell text styles are resolved; table-style inheritance is not interpreted.
    /// </summary>
    public class XSLFTableCell : IXSLFTextContainer
    {
        private const string A = XSLFXmlHelper.DrawingNamespace;
        private static readonly string[] Children = { "txBody", "tcPr", "extLst" };
        private readonly XmlElement xml;

        internal XSLFTableCell(XmlElement xml, XSLFTableRow row)
        {
            this.xml = xml;
            ParentRow = row;
        }

        /// <summary>
        /// Gets the containing physical row.
        /// </summary>
        public XSLFTableRow ParentRow { get; }

        /// <summary>
        /// Gets the containing table.
        /// </summary>
        public XSLFTable Table => ParentRow.Table;

        /// <summary>
        /// Gets actual cell text; merge anchors, styles and placeholder prompts are not substituted.
        /// </summary>
        public string Text => XSLFTextHelper.GetText(this);

        /// <summary>
        /// Gets a read-only snapshot of explicit cell paragraphs.
        /// </summary>
        public IReadOnlyList<XSLFTextParagraph> TextParagraphs => XSLFTextHelper.GetParagraphs(this);

        /// <summary>
        /// Gets the declared row span, or the schema default of one.
        /// </summary>
        public int RowSpan => Span("rowSpan");

        /// <summary>
        /// Gets the declared grid span, or the schema default of one.
        /// </summary>
        public int GridSpan => Span("gridSpan");

        /// <summary>
        /// Gets whether this cell declares a merge span or a horizontal/vertical continuation.
        /// Merge consistency and visual layout are not resolved.
        /// </summary>
        public bool IsMerged
        {
            get
            {
                var rowSpan = RowSpan;
                var gridSpan = GridSpan;
                var horizontal = Merge("hMerge");
                var vertical = Merge("vMerge");
                return rowSpan != 1 || gridSpan != 1 || horizontal || vertical;
            }
        }

        /// <summary>
        /// Replaces paragraphs and their local formatting using the shared lossless-replacement guards.
        /// Body/list styles and cell properties remain intact.
        /// </summary>
        public void SetText(string text)
        {
            XSLFTextHelper.SetText(this, text);
        }

        /// <summary>
        /// Appends an empty paragraph without modifying cell properties.
        /// </summary>
        public XSLFTextParagraph AddNewTextParagraph()
        {
            return XSLFTextHelper.AddParagraph(this);
        }

        /// <summary>
        /// Returns the containing sheet.
        /// </summary>
        public XSLFSheet GetSheet()
        {
            return Table.GetSheet();
        }

        /// <summary>
        /// Returns the live cell XML, including unsupported formatting.
        /// </summary>
        public XmlElement GetXmlObject()
        {
            return xml;
        }

        XmlElement IXSLFTextContainer.GetTextBody(bool create)
        {
            ValidateStructure();
            return XSLFTextHelper.GetBody(xml, A, create);
        }

        void IXSLFTextContainer.EnsureTextEditable()
        {
            ParentRow.EnsureEditable();
            if (!ReferenceEquals(xml.ParentNode, ParentRow.GetXmlObject()))
                throw new InvalidOperationException("The cell is detached from its row.");
        }

        internal void ValidateStructure()
        {
            XSLFTable.CheckChildren(xml, Children);
            var body = XSLFShapeHelper.RequiredChild(xml, "txBody", A);
            XSLFShapeHelper.RequiredChild(body, "bodyPr", A);
            XSLFShapeHelper.Child(body, "lstStyle", A);
            if (!body.ChildNodes.OfType<XmlElement>().Any(node => node.NamespaceURI == A && node.LocalName == "p"))
                throw new POIXMLException("A table cell text body must have at least one paragraph.");
            XSLFShapeHelper.Child(xml, "tcPr", A);
            XSLFShapeHelper.Child(xml, "extLst", A);
        }

        private int Span(string attribute)
        {
            if (!xml.HasAttribute(attribute))
                return 1;
            if (!int.TryParse(xml.GetAttribute(attribute), NumberStyles.Integer, CultureInfo.InvariantCulture,
                out var value) || value < 1)
                throw new POIXMLException("Invalid table cell span: " + attribute);
            return value;
        }

        private bool Merge(string attribute)
        {
            return xml.HasAttribute(attribute) && XSLFTextHelper.ParseBoolean(xml.GetAttribute(attribute));
        }
    }
}

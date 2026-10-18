using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A physical table row with an explicit point-based height.
    /// </summary>
    public class XSLFTableRow
    {
        private const string A = XSLFXmlHelper.DrawingNamespace;
        private static readonly string[] Children = { "tc", "extLst" };
        private readonly XmlElement xml;

        internal XSLFTableRow(XmlElement xml, XSLFTable table)
        {
            this.xml = xml;
            Table = table;
        }

        /// <summary>
        /// Gets the containing table.
        /// </summary>
        public XSLFTable Table { get; }

        /// <summary>
        /// Gets or sets the nonnegative row height in points, independently of the table anchor.
        /// </summary>
        public double Height
        {
            get { return XSLFTable.ReadDimension(xml, "h"); }
            set
            {
                EnsureEditable();
                var dimension = XSLFTable.ToDimension(value);
                xml.SetAttribute("h", dimension.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Returns a read-only snapshot of physical cells, including merged continuation cells.
        /// </summary>
        public IReadOnlyList<XSLFTableCell> GetCells()
        {
            XSLFTable.CheckChildren(xml, Children);
            XSLFShapeHelper.Child(xml, "extLst", A);
            var cells = xml.ChildNodes.OfType<XmlElement>()
                .Where(node => node.NamespaceURI == A && node.LocalName == "tc")
                .Select(node => new XSLFTableCell(node, this)).ToList();
            if (cells.Count == 0)
                throw new POIXMLException("A DrawingML table row must have at least one cell.");
            return cells.AsReadOnly();
        }

        /// <summary>
        /// Returns the live underlying row XML.
        /// </summary>
        public XmlElement GetXmlObject()
        {
            return xml;
        }

        internal void EnsureEditable()
        {
            Table.EnsureEditable();
            if (!ReferenceEquals(xml.ParentNode, Table.GetTableXml()))
                throw new InvalidOperationException("The row is detached from its table.");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Xml;
using NPOI.Util;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A fixed DrawingML table backed by preserved graphic-frame XML.
    /// Merged tables are readable but high-level editing is not supported.
    /// </summary>
    public class XSLFTable : XSLFShape
    {
        internal const string TableNamespace = "http://schemas.openxmlformats.org/drawingml/2006/table";
        internal const long MaxGridDimension = 27273042316900L;
        private const string P = XSLFShapeHelper.PresentationNamespace;
        private const string A = XSLFXmlHelper.DrawingNamespace;
        private static readonly string[] TableChildren = { "tblPr", "tblGrid", "tr", "extLst" };
        private static readonly string[] GridChildren = { "gridCol" };
        private static readonly string[] ColumnChildren = { "extLst" };

        internal XSLFTable(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }

        /// <summary>
        /// Gets the number of explicit rows.
        /// </summary>
        public int NumberOfRows => GetRows().Count;

        /// <summary>
        /// Gets the number of explicit grid columns.
        /// </summary>
        public int NumberOfColumns => GetColumns().Count;

        /// <summary>
        /// Returns a read-only snapshot of physical rows, without resolving merged-cell layout.
        /// </summary>
        public IReadOnlyList<XSLFTableRow> GetRows()
        {
            var table = GetTableXml();
            CheckChildren(table, TableChildren);
            XSLFShapeHelper.Child(table, "tblPr", A);
            XSLFShapeHelper.RequiredChild(table, "tblGrid", A);
            XSLFShapeHelper.Child(table, "extLst", A);
            var rows = table.ChildNodes.OfType<XmlElement>()
                .Where(node => node.NamespaceURI == A && node.LocalName == "tr")
                .Select(node => new XSLFTableRow(node, this)).ToList();
            if (rows.Count == 0)
                throw new POIXMLException("A DrawingML table must have at least one row.");
            return rows.AsReadOnly();
        }

        /// <summary>
        /// Returns the cell at zero-based physical row/cell indexes; merged layout is not resolved.
        /// </summary>
        public XSLFTableCell GetCell(int row, int column)
        {
            return GetRows()[row].GetCells()[column];
        }

        /// <summary>
        /// Returns a grid column width in points.
        /// </summary>
        public double GetColumnWidth(int column)
        {
            return ReadDimension(GetColumns()[column], "w");
        }

        /// <summary>
        /// Sets a grid column width in points, independently of the frame anchor.
        /// Dimensions must be finite, nonnegative DrawingML coordinates.
        /// </summary>
        public void SetColumnWidth(int column, double width)
        {
            EnsureEditable();
            var element = GetColumns()[column];
            var value = ToDimension(width);
            element.SetAttribute("w", value.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Returns a row height in points.
        /// </summary>
        public double GetRowHeight(int row)
        {
            return GetRows()[row].Height;
        }

        /// <summary>
        /// Sets a row height in points, independently of the frame anchor.
        /// </summary>
        public void SetRowHeight(int row, double height)
        {
            GetRows()[row].Height = height;
        }

        /// <inheritdoc/>
        public override SKRect GetAnchor()
        {
            return XSLFShapeHelper.ReadAnchor(GetXmlObject(), false, P);
        }

        /// <inheritdoc/>
        public override void SetAnchor(SKRect anchor)
        {
            EnsureEditable();
            XSLFShapeHelper.SetAnchor(GetXmlObject(), anchor, false, P);
        }

        internal override void EnsureEditable()
        {
            base.EnsureEditable();
            var columns = GetColumns();
            foreach (var column in columns)
                ReadDimension(column, "w");
            var rows = GetRows();
            foreach (var row in rows)
            {
                _ = row.Height;
                var cells = row.GetCells();
                if (cells.Count != columns.Count)
                    throw new NotSupportedException("Editing an irregular table grid is not supported.");
                foreach (var cell in cells)
                {
                    cell.ValidateStructure();
                    if (cell.IsMerged)
                        throw new NotSupportedException("Merged tables are read-only.");
                }
            }
        }

        internal XmlElement GetTableXml()
        {
            var graphic = XSLFShapeHelper.RequiredChild(GetXmlObject(), "graphic", A);
            var data = XSLFShapeHelper.RequiredChild(graphic, "graphicData", A);
            if (data.GetAttribute("uri") != TableNamespace ||
                data.ChildNodes.OfType<XmlElement>().Any(node => node.NamespaceURI != A || node.LocalName != "tbl"))
                throw new NotSupportedException("Opaque graphic-frame content is not interpreted as a table.");
            return XSLFShapeHelper.RequiredChild(data, "tbl", A);
        }

        internal static bool IsTable(XmlElement xml)
        {
            var graphic = XSLFShapeHelper.Child(xml, "graphic", A);
            var data = graphic == null ? null : XSLFShapeHelper.Child(graphic, "graphicData", A);
            return data != null && data.GetAttribute("uri") == TableNamespace &&
                XSLFShapeHelper.Child(data, "tbl", A) != null;
        }

        internal static XSLFTable Create(
            XSLFSheet sheet, IXSLFShapeContainer parent, XmlElement tree, int rows, int columns)
        {
            XSLFShapeHelper.EnsureAttached(sheet, tree);
            XSLFShapeHelper.ValidateContainer(tree);
            if (rows <= 0)
                throw new ArgumentOutOfRangeException(nameof(rows));
            if (columns <= 0)
                throw new ArgumentOutOfRangeException(nameof(columns));
            var id = XSLFShapeHelper.NextId(sheet).ToString(CultureInfo.InvariantCulture);
            var xml = XSLFXmlHelper.CreateElement(
                "<p:graphicFrame xmlns:p=\"" + P + "\" xmlns:a=\"" + A +
                "\"><p:nvGraphicFramePr><p:cNvPr id=\"" + id + "\" name=\"Table " + id +
                "\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" +
                "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></p:xfrm>" +
                "<a:graphic><a:graphicData uri=\"" + TableNamespace +
                "\"><a:tbl><a:tblPr/><a:tblGrid/></a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
            var table = xml["graphic", A]["graphicData", A]["tbl", A];
            for (var column = 0; column < columns; column++)
            {
                var element = xml.OwnerDocument.CreateElement("a", "gridCol", A);
                element.SetAttribute("w", "0");
                table["tblGrid", A].AppendChild(element);
            }
            var prototype = XSLFXmlHelper.CreateElement("<a:tc xmlns:a=\"" + A +
                "\"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>");
            for (var row = 0; row < rows; row++)
            {
                var element = xml.OwnerDocument.CreateElement("a", "tr", A);
                element.SetAttribute("h", "0");
                for (var column = 0; column < columns; column++)
                    element.AppendChild(xml.OwnerDocument.ImportNode(prototype, true));
                table.AppendChild(element);
            }
            return new XSLFTable(XSLFShapeHelper.Append(tree, xml), sheet, parent);
        }

        internal static double ReadDimension(XmlElement element, string attribute)
        {
            if (!long.TryParse(element.GetAttribute(attribute), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var value) || value < 0 || value > MaxGridDimension)
                throw new POIXMLException("Missing or invalid table grid dimension: " + attribute);
            return Units.ToPoints(value);
        }

        internal static long ToDimension(double points)
        {
            var value = Math.Round(points * Units.EMU_PER_POINT);
            if (double.IsNaN(points) || double.IsInfinity(points) || points < 0 || value > MaxGridDimension)
                throw new ArgumentOutOfRangeException(nameof(points),
                    "Table grid dimensions must be finite nonnegative DrawingML coordinates.");
            return (long)value;
        }

        internal static void CheckChildren(XmlElement parent, string[] names)
        {
            if (parent.ChildNodes.OfType<XmlElement>().Any(node => node.NamespaceURI != A || !names.Contains(node.LocalName)))
                throw new NotSupportedException("Opaque table structure is not interpreted or edited.");
        }

        private ReadOnlyCollection<XmlElement> GetColumns()
        {
            var grid = XSLFShapeHelper.RequiredChild(GetTableXml(), "tblGrid", A);
            CheckChildren(grid, GridChildren);
            var columns = grid.ChildNodes.OfType<XmlElement>().ToList();
            if (columns.Count == 0)
                throw new POIXMLException("A DrawingML table must have at least one grid column.");
            foreach (var column in columns)
            {
                CheckChildren(column, ColumnChildren);
                XSLFShapeHelper.Child(column, "extLst", A);
            }
            return columns.AsReadOnly();
        }
    }
}

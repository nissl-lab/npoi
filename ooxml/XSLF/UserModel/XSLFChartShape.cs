using System;
using System.Globalization;
using System.Linq;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// An editable graphic frame referencing a read-only classic DrawingML chart.
    /// Chart contents, embedded workbooks, styles and external data are not modified.
    /// </summary>
    public class XSLFChartShape : XSLFShape
    {
        private const string P = CT_Presentation.Namespace;
        private const string A = XSLFXmlHelper.DrawingNamespace;
        private const string C = XSLFChart.ChartNamespace;
        private const string R = CT_Presentation.RelationshipsNamespace;

        internal XSLFChartShape(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }

        /// <summary>
        /// Resolves the referenced classic chart in this presentation.
        /// External chart targets and malformed references fail explicitly; no URI is fetched.
        /// </summary>
        public XSLFChart GetChart()
        {
            var presentation = GetSheet().GetSlideShow();
            presentation.EnsureOpen();
            var reference = GetReference();
            var id = reference.GetAttribute("id", R);
            var source = GetSheet().GetPackagePart();
            var relationship = string.IsNullOrEmpty(id) ? null : source.GetRelationship(id);
            if (relationship == null || relationship.TargetMode != TargetMode.Internal ||
                relationship.RelationshipType != XSLFRelation.CHART.Relation)
                throw new POIXMLException("Missing or invalid chart relationship: " + id);
            if (relationship.TargetUri.OriginalString.Contains('#'))
                throw new NotSupportedException("Chart relationship fragments are not supported.");
            return presentation.GetChart(source.GetRelatedPart(relationship));
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
            GetChart().GetXmlSnapshot();
        }

        internal static bool IsChart(XmlElement xml)
        {
            var graphic = XSLFShapeHelper.Child(xml, "graphic", A);
            var data = graphic == null ? null : XSLFShapeHelper.Child(graphic, "graphicData", A);
            return data != null && data.GetAttribute("uri") == C &&
                XSLFShapeHelper.Child(data, "chart", C) != null;
        }

        internal static XSLFChartShape Create(
            XSLFSheet sheet, IXSLFShapeContainer parent, XmlElement tree, XSLFChart chart)
        {
            XSLFShapeHelper.EnsureAttached(sheet, tree);
            XSLFShapeHelper.ValidateContainer(tree);
            if (chart == null)
                throw new ArgumentNullException(nameof(chart));
            if (!ReferenceEquals(chart.GetSlideShow(), sheet.GetSlideShow()))
                throw new ArgumentException("The chart must belong to this presentation.", nameof(chart));
            chart.GetXmlSnapshot();
            var id = XSLFShapeHelper.NextId(sheet).ToString(CultureInfo.InvariantCulture);
            var xml = XSLFXmlHelper.CreateElement(
                "<p:graphicFrame xmlns:p=\"" + P + "\" xmlns:a=\"" + A + "\" xmlns:c=\"" + C +
                "\" xmlns:r=\"" + R + "\"><p:nvGraphicFramePr><p:cNvPr id=\"" + id + "\" name=\"Chart " + id +
                "\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" +
                "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></p:xfrm>" +
                "<a:graphic><a:graphicData uri=\"" + C +
                "\"><c:chart r:id=\"\"/></a:graphicData></a:graphic></p:graphicFrame>");
            var relationship = sheet.GetOrAddRelation(chart, XSLFRelation.CHART);
            xml["graphic", A]["graphicData", A]["chart", C].SetAttribute("id", R, relationship.Id);
            return new XSLFChartShape(XSLFShapeHelper.Append(tree, xml), sheet, parent);
        }

        private XmlElement GetReference()
        {
            var graphic = XSLFShapeHelper.RequiredChild(GetXmlObject(), "graphic", A);
            var data = XSLFShapeHelper.RequiredChild(graphic, "graphicData", A);
            if (data.GetAttribute("uri") != C || data.ChildNodes.OfType<XmlElement>()
                .Any(node => node.NamespaceURI != C || node.LocalName != "chart"))
                throw new NotSupportedException("Opaque graphic-frame content is not interpreted as a chart.");
            var reference = XSLFShapeHelper.RequiredChild(data, "chart", C);
            if (reference.ChildNodes.OfType<XmlElement>().Any())
                throw new NotSupportedException("Opaque chart-reference content is not interpreted.");
            return reference;
        }
    }
}

using System;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.Util;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A byte-preserving, read-only classic DrawingML chart part.
    /// XML and schema-model snapshots are detached; their changes are not written back.
    /// Chart authoring, data/workbook synchronization and rendering are not supported.
    /// </summary>
    public class XSLFChart : POIXMLDocumentPart
    {
        internal const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        internal XSLFChart(POIXMLDocumentPart parent, PackagePart part) : base(parent, part)
        {
            Committed = true;
        }

        /// <summary>
        /// Gets a fresh copy of the original chart bytes without loading related workbooks or external data.
        /// </summary>
        public byte[] Data
        {
            get
            {
                EnsureAttached();
                using var stream = GetPackagePart().GetInputStream();
                return IOUtils.ToByteArray(stream);
            }
        }

        /// <summary>
        /// Returns a complete detached XML snapshot, preserving unknown elements and attributes.
        /// Snapshot changes do not modify the chart. DTDs and external entity resolution are prohibited.
        /// </summary>
        public XmlElement GetXmlSnapshot()
        {
            EnsureAttached();
            var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
            using var stream = GetPackagePart().GetInputStream();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            try
            {
                document.Load(reader);
                var root = document.DocumentElement;
                if (root == null || root.LocalName != "chartSpace" || root.NamespaceURI != ChartNamespace)
                    throw new POIXMLException("Invalid DrawingML chart root: " + GetPackagePart().PartName.Name);
                var chart = XSLFShapeHelper.RequiredChild(root, "chart", ChartNamespace);
                XSLFShapeHelper.RequiredChild(chart, "plotArea", ChartNamespace);
                return root;
            }
            catch (XmlException e)
            {
                throw new POIXMLException("Invalid DrawingML chart XML: " + GetPackagePart().PartName.Name, e);
            }
        }

        /// <summary>
        /// Returns an independent snapshot using the existing partial chart schema model.
        /// Unsupported XML may not be represented by that model; use GetXmlSnapshot for the full XML.
        /// Model changes do not modify the chart or load/recalculate related workbooks.
        /// </summary>
        public CT_ChartSpace GetChartSpaceSnapshot()
        {
            return CT_ChartSpace.Parse(GetXmlSnapshot(), NamespaceManager);
        }

        internal XMLSlideShow GetSlideShow()
        {
            for (var parent = GetParent(); parent != null; parent = parent.GetParent())
                if (parent is XMLSlideShow presentation)
                    return presentation;
            throw new POIXMLException("The chart is not attached to a presentation.");
        }

        internal void EnsureAttached()
        {
            var presentation = GetSlideShow();
            presentation.EnsureOpen();
            var part = GetPackagePart();
            if (!ReferenceEquals(part.Package, presentation.Package) ||
                !ReferenceEquals(presentation.Package.GetPart(part.PartName), part))
                throw new InvalidOperationException("The chart part is no longer attached to this presentation.");
            if (!string.Equals(part.ContentType, XSLFRelation.CHART.ContentType, StringComparison.OrdinalIgnoreCase))
                throw new POIXMLException("The part is no longer a classic DrawingML chart.");
        }

        protected internal override void PrepareForCommit()
        {
            // The partial model is never serialized; preserve the original chart and its dependencies.
        }
    }
}

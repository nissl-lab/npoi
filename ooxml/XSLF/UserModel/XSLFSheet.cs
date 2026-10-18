using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Shared shapes, package navigation and XML persistence for slides, layouts, and masters.
    /// </summary>
    public abstract class XSLFSheet : POIXMLDocumentPart, IXSLFShapeContainer
    {
        private byte[] savedXml;

        protected XSLFSheet()
        {
        }

        protected XSLFSheet(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
        }

        /// <summary>
        /// Returns the presentation containing this sheet.
        /// </summary>
        public XMLSlideShow GetSlideShow()
        {
            var parent = GetParent();
            while (parent != null)
            {
                if (parent is XMLSlideShow slideShow)
                {
                    return slideShow;
                }
                parent = parent.GetParent();
            }
            throw new POIXMLException("The sheet is not attached to a presentation.");
        }

        /// <summary>
        /// Returns the layout or master used by this sheet, or null for a slide master.
        /// </summary>
        public abstract XSLFSheet GetMasterSheet();

        /// <summary>
        /// Returns the inherited theme, or null if the master has no theme.
        /// </summary>
        public abstract XSLFTheme GetTheme();

        /// <inheritdoc/>
        public IReadOnlyList<XSLFShape> GetShapes()
        {
            return XSLFShapeHelper.GetShapes(this, this, GetShapeTree());
        }

        /// <inheritdoc/>
        public XSLFSimpleShape CreateAutoShape(ST_ShapeType shapeType = ST_ShapeType.rect)
        {
            EnsureShapesWritable();
            return XSLFShapeHelper.CreateAutoShape(this, this, GetShapeTree(), shapeType);
        }

        /// <inheritdoc/>
        public XSLFGroupShape CreateGroup()
        {
            EnsureShapesWritable();
            return XSLFShapeHelper.CreateGroup(this, this, GetShapeTree());
        }

        protected abstract XmlElement GetCommonSlideData();

        internal XmlElement GetShapeTree()
        {
            return XSLFShapeHelper.RequiredChild(GetCommonSlideData(), "spTree", XSLFShapeHelper.PresentationNamespace);
        }

        internal void EnsureShapesWritable()
        {
            var presentation = GetSlideShow();
            presentation.EnsureWritable();
            var part = GetPackagePart();
            if (Committed || part == null || !ReferenceEquals(part.Package, presentation.Package) ||
                !ReferenceEquals(presentation.Package.GetPart(part.PartName), part))
            {
                throw new InvalidOperationException("Shape changes to a detached or committed sheet would not be saved.");
            }
        }

        protected void InitializeXml()
        {
            using var output = new MemoryStream();
            WriteXml(output);
            savedXml = output.ToArray();
        }

        protected abstract void WriteXml(Stream stream);

        protected internal override void PrepareForCommit()
        {
            // Keep the original part bytes until serialization confirms that the model changed.
        }

        protected internal override void Commit()
        {
            using var output = new MemoryStream();
            WriteXml(output);
            var xml = output.ToArray();
            if (savedXml != null && savedXml.SequenceEqual(xml))
            {
                return;
            }
            var part = GetPackagePart();
            part.Clear();
            using var stream = part.GetOutputStream();
            stream.Write(xml, 0, xml.Length);
            savedXml = xml;
        }

        internal T GetRequiredRelation<T>(XSLFRelation descriptor) where T : POIXMLDocumentPart
        {
            var relationships = GetPackagePart().GetRelationshipsByType(descriptor.Relation);
            if (relationships.Size != 1)
            {
                throw new POIXMLException("Expected one " + descriptor.Relation + " relationship in " +
                    GetPackagePart().PartName.Name);
            }
            return XSLFXmlHelper.GetReferencedPart<T>(this, relationships.GetRelationship(0).Id, descriptor);
        }
    }
}

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

        /// <inheritdoc/>
        public XSLFTextBox CreateTextBox()
        {
            EnsureShapesWritable();
            return XSLFShapeHelper.CreateTextBox(this, this, GetShapeTree());
        }

        /// <inheritdoc/>
        public XSLFPictureShape CreatePicture(XSLFPictureData pictureData)
        {
            return XSLFPictureShape.Create(this, this, GetShapeTree(), pictureData);
        }

        /// <inheritdoc/>
        public XSLFTable CreateTable(int rows, int columns)
        {
            return XSLFTable.Create(this, this, GetShapeTree(), rows, columns);
        }

        /// <inheritdoc/>
        public XSLFChartShape CreateChart(XSLFChart chart)
        {
            return XSLFChartShape.Create(this, this, GetShapeTree(), chart);
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

        internal PackageRelationship GetOrAddRelation(POIXMLDocumentPart child, XSLFRelation descriptor)
        {
            EnsureShapesWritable();
            var source = GetPackagePart();
            var target = child.GetPackagePart();
            if (!ReferenceEquals(target.Package, source.Package) ||
                !ReferenceEquals(source.Package.GetPart(target.PartName), target))
                throw new ArgumentException("The related part must belong to this presentation.", nameof(child));
            PackageRelationship relationship = null;
            foreach (var candidate in source.GetRelationshipsByType(descriptor.Relation))
            {
                if (candidate.TargetMode == TargetMode.Internal &&
                    string.Equals(candidate.TargetUri.OriginalString, target.PartName.Name, StringComparison.OrdinalIgnoreCase))
                {
                    relationship = candidate;
                    break;
                }
            }
            if (relationship != null)
            {
                var cached = GetRelationById(relationship.Id);
                if (cached != null && !ReferenceEquals(cached.GetPackagePart(), target))
                    throw new POIXMLException("The cached document relation does not match its package target.");
            }
            // AddRelation's target-only lookup can reuse a relationship of the wrong type.
            relationship ??= source.AddRelationship(target.PartName, TargetMode.Internal, descriptor.Relation);
            if (GetRelationById(relationship.Id) == null)
                AddRelation(relationship, child);
            return relationship;
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

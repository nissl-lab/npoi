using System;
using System.Globalization;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// An embedded or externally linked picture, preserving its existing DrawingML.
    /// Only embedded PNG/JPEG pictures support high-level name and anchor changes.
    /// </summary>
    public class XSLFPictureShape : XSLFShape
    {
        private const string P = CT_Presentation.Namespace;
        private const string R = CT_Presentation.RelationshipsNamespace;

        internal XSLFPictureShape(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }

        /// <summary>
        /// Gets whether the primary blip has an external link, even when an embedded cache also exists.
        /// </summary>
        public bool IsExternalLinkedPicture => Blip.HasAttribute("link", R);

        /// <summary>
        /// Returns the primary embedded image, or null for a validated external-only picture.
        /// SVG alternatives, effects and other extension images are not resolved.
        /// </summary>
        public XSLFPictureData GetPictureData()
        {
            if (!Blip.HasAttribute("embed", R))
            {
                if (GetPictureLink() != null)
                {
                    return null;
                }
                throw new POIXMLException("The picture has no embedded or linked image.");
            }
            var relation = GetImageRelationship("embed", TargetMode.Internal);
            if (relation.TargetUri.OriginalString.Contains('#'))
            {
                throw new NotSupportedException("Image relationship fragments are not supported.");
            }
            var part = GetSheet().GetPackagePart().GetRelatedPart(relation);
            return GetSheet().GetSlideShow().GetPictureData(part);
        }

        /// <summary>
        /// Returns the external image URI, including for pictures with an embedded cache, or null.
        /// The URI is never fetched.
        /// </summary>
        public Uri GetPictureLink()
        {
            return Blip.HasAttribute("link", R)
                ? GetImageRelationship("link", TargetMode.External).TargetUri : null;
        }

        /// <inheritdoc/>
        public override SKRect GetAnchor()
        {
            return XSLFShapeHelper.ReadAnchor(Properties, false);
        }

        /// <inheritdoc/>
        public override void SetAnchor(SKRect anchor)
        {
            EnsureEditable();
            XSLFShapeHelper.SetAnchor(Properties, anchor, false);
        }

        internal override void EnsureEditable()
        {
            base.EnsureEditable();
            if (IsExternalLinkedPicture)
            {
                GetPictureLink();
                throw new NotSupportedException("Externally linked pictures are read-only.");
            }
            XSLFPictureData.GetDescriptor(GetPictureData().PictureType);
        }

        internal static XSLFPictureShape Create(
            XSLFSheet sheet, IXSLFShapeContainer parent, XmlElement tree, XSLFPictureData data)
        {
            XSLFShapeHelper.EnsureAttached(sheet, tree);
            XSLFShapeHelper.ValidateContainer(tree);
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }
            if (!ReferenceEquals(data.GetSlideShow(), sheet.GetSlideShow()))
            {
                throw new ArgumentException("The picture data must belong to this presentation.", nameof(data));
            }
            data.GetImageDimensionInPixels();
            var id = XSLFShapeHelper.NextId(sheet).ToString(CultureInfo.InvariantCulture);
            var xml = XSLFXmlHelper.CreateElement(
                "<p:pic xmlns:p=\"" + P + "\" xmlns:a=\"" + XSLFXmlHelper.DrawingNamespace +
                "\" xmlns:r=\"" + R + "\"><p:nvPicPr><p:cNvPr id=\"" + id + "\" name=\"Picture " + id +
                "\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>" +
                "<p:blipFill><a:blip r:embed=\"\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>" +
                "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
            var source = sheet.GetPackagePart();
            PackageRelationship relationship = null;
            foreach (var candidate in source.GetRelationshipsByType(PackageRelationshipTypes.IMAGE_PART))
            {
                if (candidate.TargetMode == TargetMode.Internal &&
                    string.Equals(candidate.TargetUri.OriginalString, data.GetPackagePart().PartName.Name,
                        StringComparison.OrdinalIgnoreCase))
                {
                    relationship = candidate;
                    break;
                }
            }
            // AddRelation can reuse a non-image relationship to the same part; match type and target explicitly.
            relationship ??= source.AddRelationship(
                data.GetPackagePart().PartName, TargetMode.Internal, PackageRelationshipTypes.IMAGE_PART);
            if (sheet.GetRelationById(relationship.Id) == null)
            {
                sheet.AddRelation(relationship, data);
            }
            xml["blipFill", P]["blip", XSLFXmlHelper.DrawingNamespace].SetAttribute("embed", R, relationship.Id);
            return new XSLFPictureShape(XSLFShapeHelper.Append(tree, xml), sheet, parent);
        }

        private PackageRelationship GetImageRelationship(string attribute, TargetMode mode)
        {
            var id = Blip.GetAttribute(attribute, R);
            var relationship = GetSheet().GetPackagePart().GetRelationship(id);
            if (string.IsNullOrEmpty(id) || relationship == null ||
                relationship.RelationshipType != PackageRelationshipTypes.IMAGE_PART ||
                relationship.TargetMode != mode || relationship.TargetUri.OriginalString.Length == 0)
            {
                throw new POIXMLException("Missing or invalid picture " + attribute + " relationship.");
            }
            return relationship;
        }

        private XmlElement Blip => XSLFShapeHelper.RequiredChild(
            XSLFShapeHelper.RequiredChild(GetXmlObject(), "blipFill", P), "blip", XSLFXmlHelper.DrawingNamespace);

        private XmlElement Properties => XSLFShapeHelper.RequiredChild(GetXmlObject(), "spPr", P);
    }
}

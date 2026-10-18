using System;
using System.Collections.Generic;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Describes the package relationships used by the initial XSLF implementation.
    /// </summary>
    public class XSLFRelation : POIXMLRelation
    {
        private static readonly Dictionary<string, XSLFRelation> Relations =
            new Dictionary<string, XSLFRelation>(StringComparer.Ordinal);

        public static readonly XSLFRelation MAIN = new XSLFRelation(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            PackageRelationshipTypes.CORE_DOCUMENT, "/ppt/presentation.xml", null);

        public static readonly XSLFRelation SLIDE = new XSLFRelation(
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            CT_Presentation.RelationshipsNamespace + "/slide", "/ppt/slides/slide#.xml", typeof(XSLFSlide));

        public static readonly XSLFRelation SLIDE_LAYOUT = new XSLFRelation(
            "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
            CT_Presentation.RelationshipsNamespace + "/slideLayout", "/ppt/slideLayouts/slideLayout#.xml",
            typeof(XSLFSlideLayout));

        public static readonly XSLFRelation SLIDE_MASTER = new XSLFRelation(
            "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
            CT_Presentation.RelationshipsNamespace + "/slideMaster", "/ppt/slideMasters/slideMaster#.xml",
            typeof(XSLFSlideMaster));

        public static readonly XSLFRelation THEME = new XSLFRelation(
            "application/vnd.openxmlformats-officedocument.theme+xml",
            CT_Presentation.RelationshipsNamespace + "/theme", "/ppt/theme/theme#.xml", typeof(XSLFTheme));

        public static readonly XSLFRelation IMAGE_PNG = new XSLFRelation(
            "image/png", PackageRelationshipTypes.IMAGE_PART, "/ppt/media/image#.png", typeof(XSLFPictureData));

        public static readonly XSLFRelation IMAGE_JPEG = new XSLFRelation(
            "image/jpeg", PackageRelationshipTypes.IMAGE_PART, "/ppt/media/image#.jpeg", typeof(XSLFPictureData));

        public static readonly XSLFRelation CHART = new XSLFRelation(
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            CT_Presentation.RelationshipsNamespace + "/chart", "/ppt/charts/chart#.xml", typeof(XSLFChart));

        private XSLFRelation(string contentType, string relationship, string defaultName, Type type)
            : base(contentType, relationship, defaultName, type)
        {
            if (!Relations.ContainsKey(relationship))
            {
                Relations.Add(relationship, this);
            }
        }

        /// <summary>
        /// Returns the descriptor for a supported relationship, or null for an unknown relationship.
        /// </summary>
        public static XSLFRelation GetInstance(string relationshipType)
        {
            return Relations.TryGetValue(relationshipType, out var relation) ? relation : null;
        }
    }
}

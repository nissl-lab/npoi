using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.Util;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.Util;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Reads, writes, and creates transitional PPTX presentations and blank slides.
    /// High-level shape editing and rendering are not yet supported.
    /// </summary>
    public class XMLSlideShow : POIXMLDocument, IDisposable
    {
        private readonly List<XSLFSlide> slides = new List<XSLFSlide>();
        private CT_Presentation presentation;

        /// <summary>
        /// Creates an empty presentation with a blank layout, slide master, and theme.
        /// </summary>
        public XMLSlideShow() : this(NewPackage())
        {
        }

        /// <summary>
        /// Opens a PPTX stream using the existing OOXML package stream-ownership rules.
        /// </summary>
        public XMLSlideShow(Stream stream) : this(PackageHelper.Open(stream))
        {
        }

        /// <summary>
        /// Opens a PPTX package. Closing this presentation also closes the package.
        /// </summary>
        public XMLSlideShow(OPCPackage package) : base(ValidatePackage(package))
        {
            Load(XSLFFactory.GetInstance());
        }

        /// <summary>
        /// Returns the slides in presentation order as a read-only collection.
        /// </summary>
        public IReadOnlyList<XSLFSlide> GetSlides()
        {
            return slides.AsReadOnly();
        }

        /// <summary>
        /// Returns the underlying partial PresentationML model.
        /// </summary>
        public CT_Presentation GetCTPresentation()
        {
            return presentation;
        }

        /// <summary>
        /// Appends a blank slide using an existing blank layout.
        /// </summary>
        /// <exception cref="NotSupportedException">The presentation has no blank layout.</exception>
        public XSLFSlide CreateSlide()
        {
#if NET8_0_OR_GREATER
            ObjectDisposedException.ThrowIf(Package == null, this);
#else
            if (Package == null)
            {
                throw new ObjectDisposedException(nameof(XMLSlideShow));
            }
#endif
            if (Package.GetPackageAccess() == PackageAccess.READ)
            {
                throw new InvalidOperationException("Cannot create slides in a read-only presentation.");
            }

            var layout = FindBlankLayout();
            var entries = presentation.sldIdLst?.sldId;
            var maxId = entries == null || entries.Count == 0 ? 255U : entries.Max(entry => entry.id);
            if (maxId >= int.MaxValue)
            {
                throw new InvalidOperationException("No more PresentationML slide identifiers are available.");
            }
            var partNumber = GetNextPartNumber(XSLFRelation.SLIDE, 1);
            if (partNumber < 0)
            {
                throw new InvalidOperationException("No free slide part name is available.");
            }

            var slide = (XSLFSlide)CreateRelationship(XSLFRelation.SLIDE, XSLFFactory.GetInstance(), partNumber);
            slide.AddRelation(null, XSLFRelation.SLIDE_LAYOUT, layout);
            if (presentation.sldIdLst == null)
            {
                presentation.sldIdLst = new CT_SlideIdList();
            }
            presentation.sldIdLst.sldId.Add(new CT_SlideIdListEntry
            {
                id = maxId + 1,
                relationshipId = GetRelationId(slide)
            });
            slides.Add(slide);
            return slide;
        }

        /// <summary>
        /// Returns package parts stored in the presentation's embeddings directory.
        /// </summary>
        public override List<PackagePart> GetAllEmbedds()
        {
            return Package.GetParts()
                .Where(part => part.PartName.Name.StartsWith("/ppt/embeddings/", StringComparison.Ordinal))
                .ToList();
        }

        internal override void OnDocumentRead()
        {
            presentation = XSLFXmlHelper.Read<CT_Presentation>(GetPackagePart());
            var identifiers = new HashSet<uint>();
            var slideParts = new HashSet<XSLFSlide>();
            foreach (var entry in presentation.sldIdLst?.sldId ?? new List<CT_SlideIdListEntry>())
            {
                if (entry.id < 256 || entry.id > int.MaxValue || !identifiers.Add(entry.id))
                {
                    throw new POIXMLException("Invalid or duplicate presentation slide identifier: " + entry.id);
                }
                var relationship = string.IsNullOrEmpty(entry.relationshipId)
                    ? null : GetPackagePart().GetRelationship(entry.relationshipId);
                if (relationship == null || relationship.TargetMode != TargetMode.Internal ||
                    relationship.RelationshipType != XSLFRelation.SLIDE.Relation ||
                    !(GetRelationById(entry.relationshipId) is XSLFSlide slide))
                {
                    throw new POIXMLException("Invalid slide relationship: " + entry.relationshipId);
                }
                if (!slideParts.Add(slide))
                {
                    throw new POIXMLException("A slide is referenced more than once: " + entry.relationshipId);
                }
                slides.Add(slide);
            }
        }

        protected internal override void Commit()
        {
            XSLFXmlHelper.Write(GetPackagePart(), presentation);
        }

        /// <summary>
        /// Closes the underlying package.
        /// </summary>
        public void Dispose()
        {
            Close();
        }

        private POIXMLDocumentPart FindBlankLayout()
        {
            foreach (var master in GetRelations().Where(part =>
                part.GetPackagePart().ContentType == XSLFRelation.SLIDE_MASTER.ContentType))
            {
                foreach (var layout in master.GetRelations().Where(part =>
                    part.GetPackagePart().ContentType == XSLFRelation.SLIDE_LAYOUT.ContentType))
                {
                    var document = new XmlDocument { XmlResolver = null };
                    using var stream = layout.GetPackagePart().GetInputStream();
                    XmlHelper.LoadXmlSafe(document, stream);
                    var root = document.DocumentElement;
                    if (root != null && root.LocalName == "sldLayout" &&
                        root.NamespaceURI == CT_Presentation.Namespace && root.GetAttribute("type") == "blank")
                    {
                        return layout;
                    }
                }
            }
            throw new NotSupportedException("Creating a blank slide requires a presentation with a blank slide layout.");
        }

        private static OPCPackage ValidatePackage(OPCPackage package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }
            var relationships = package.GetRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
            if (relationships.Size != 1 || relationships.GetRelationship(0).TargetMode != TargetMode.Internal)
            {
                throw new POIXMLException("A PPTX package must have exactly one internal office document relationship.");
            }
            var part = package.GetPart(relationships.GetRelationship(0));
            if (part == null || part.ContentType != XSLFRelation.MAIN.ContentType)
            {
                throw new POIXMLException("The package is not a supported transitional PPTX presentation.");
            }
            return package;
        }

        private static OPCPackage NewPackage()
        {
            var package = OPCPackage.Create(new MemoryStream());
            var main = package.CreatePart(
                PackagingUriHelper.CreatePartName(XSLFRelation.MAIN.DefaultFileName), XSLFRelation.MAIN.ContentType);
            package.AddRelationship(main.PartName, TargetMode.Internal, XSLFRelation.MAIN.Relation);
            var master = CreateResourcePart(package, XSLFRelation.SLIDE_MASTER, "slideMaster.xml");
            var layout = CreateResourcePart(package, XSLFRelation.SLIDE_LAYOUT, "slideLayout.xml");
            var theme = CreateResourcePart(package, XSLFRelation.THEME, "theme.xml");
            main.AddRelationship(master.PartName, TargetMode.Internal, XSLFRelation.SLIDE_MASTER.Relation, "rId1");
            master.AddRelationship(layout.PartName, TargetMode.Internal, XSLFRelation.SLIDE_LAYOUT.Relation, "rId1");
            master.AddRelationship(theme.PartName, TargetMode.Internal, XSLFRelation.THEME.Relation, "rId2");
            layout.AddRelationship(master.PartName, TargetMode.Internal, XSLFRelation.SLIDE_MASTER.Relation);

            var document = new CT_Presentation
            {
                Namespaces = XSLFXmlHelper.CreateNamespaces(),
                sldMasterIdLst = XSLFXmlHelper.CreateElement(
                    "<p:sldMasterIdLst xmlns:p=\"" + CT_Presentation.Namespace + "\" xmlns:r=\"" +
                    CT_Presentation.RelationshipsNamespace +
                    "\"><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>"),
                sldSz = XSLFXmlHelper.CreateElement(
                    "<p:sldSz xmlns:p=\"" + CT_Presentation.Namespace +
                    "\" cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"/>"),
                notesSz = XSLFXmlHelper.CreateElement(
                    "<p:notesSz xmlns:p=\"" + CT_Presentation.Namespace + "\" cx=\"6858000\" cy=\"9144000\"/>")
            };
            XSLFXmlHelper.Write(main, document);
            return package;
        }

        private static PackagePart CreateResourcePart(OPCPackage package, XSLFRelation relation, string name)
        {
            var part = package.CreatePart(PackagingUriHelper.CreatePartName(relation.GetFileName(1)), relation.ContentType);
            using var resource = typeof(XMLSlideShow).Assembly.GetManifestResourceStream("NPOI.XSLF.Resources." + name);
            if (resource == null)
            {
                throw new POIXMLException("Missing XSLF resource: " + name);
            }
            using var output = part.GetOutputStream();
            resource.CopyTo(output);
            return part;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.Util;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Reads, writes, and creates transitional PPTX presentations and layout-based slides.
    /// High-level shape editing and rendering are not yet supported.
    /// </summary>
    public class XMLSlideShow : POIXMLDocument, IDisposable
    {
        private const int MinSlideSize = 914400;
        private const int MaxSlideSize = 51206400;
        private const string SectionNamespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";

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
        /// Returns referenced slide masters in presentation order as a read-only collection.
        /// </summary>
        public IReadOnlyList<XSLFSlideMaster> GetSlideMasters()
        {
            return XSLFXmlHelper.GetReferencedParts<XSLFSlideMaster>(
                this, presentation.sldMasterIdLst, "sldMasterId", XSLFRelation.SLIDE_MASTER).AsReadOnly();
        }

        /// <summary>
        /// Returns the underlying partial PresentationML model.
        /// </summary>
        public CT_Presentation GetCTPresentation()
        {
            return presentation;
        }

        /// <summary>
        /// Appends a blank slide using an existing blank layout, reserving IDs in all preserved compatibility branches.
        /// </summary>
        /// <exception cref="NotSupportedException">The presentation has no blank layout or has an opaque wrapped slide list.</exception>
        /// <exception cref="POIXMLException">A preserved slide identifier is invalid.</exception>
        /// <exception cref="InvalidOperationException">The presentation is read-only or no slide identifier or part name is available.</exception>
        public XSLFSlide CreateSlide()
        {
            EnsureCanCreateSlide();
            return CreateSlide(FindBlankLayout());
        }

        /// <summary>
        /// Appends a slide referencing a layout in this presentation.
        /// Placeholder shapes are not copied from the layout in this implementation.
        /// </summary>
        /// <exception cref="NotSupportedException">The presentation has an opaque wrapped slide list.</exception>
        /// <exception cref="POIXMLException">A referenced part or preserved slide identifier is invalid.</exception>
        /// <exception cref="InvalidOperationException">The presentation is read-only or no slide identifier or part name is available.</exception>
        public XSLFSlide CreateSlide(XSLFSlideLayout layout)
        {
            EnsureCanCreateSlide();
            if (layout == null)
            {
                throw new ArgumentNullException(nameof(layout));
            }
            if (!ReferenceEquals(layout.GetPackagePart().Package, Package) ||
                !ReferenceEquals(layout.GetSlideShow(), this))
            {
                throw new ArgumentException("The slide layout must belong to this presentation.", nameof(layout));
            }
            var master = layout.GetSlideMaster();
            if (!GetSlideMasters().Contains(master) || !master.GetSlideLayouts().Contains(layout))
            {
                throw new ArgumentException("The layout must be listed by a referenced slide master.", nameof(layout));
            }
            var entries = presentation.sldIdLst?.sldId;
            var maxId = entries == null || entries.Count == 0 ? 255U : entries.Max(entry => entry.id);
            foreach (var element in presentation.sldIdLst?.AnyElements ?? Array.Empty<XmlElement>())
            {
                // Every compatibility branch must remain collision-free, not just the selected branch.
                foreach (XmlElement reference in element.GetElementsByTagName("sldId", CT_Presentation.Namespace))
                {
                    var value = reference.GetAttribute("id");
                    if (!uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ||
                        id < 256 || id > int.MaxValue)
                    {
                        throw new POIXMLException("Invalid preserved presentation slide identifier: " + value);
                    }
                    maxId = Math.Max(maxId, id);
                }
            }
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
        /// Moves a slide to a zero-based position, preserving slide and relationship identifiers.
        /// Presentations with sections or opaque presentation extensions cannot be reordered yet.
        /// </summary>
        public void SetSlideOrder(XSLFSlide slide, int newIndex)
        {
            EnsureWritable();
            if (slide == null)
            {
                throw new ArgumentNullException(nameof(slide));
            }
            var oldIndex = slides.IndexOf(slide);
            if (oldIndex < 0)
            {
                throw new ArgumentException("The slide must belong to this presentation.", nameof(slide));
            }
            if (newIndex < 0 || newIndex >= slides.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(newIndex));
            }
            var entries = GetConsistentSlideEntries();
            if (oldIndex == newIndex)
            {
                return;
            }
            if ((presentation.AnyElements ?? Array.Empty<XmlElement>()).Any(root =>
                    root.NamespaceURI != CT_Presentation.Namespace) ||
                GetPresentationElements().Any(element =>
                (element.LocalName == "ext" && element.NamespaceURI == CT_Presentation.Namespace &&
                    element.ChildNodes.OfType<XmlElement>().Any()) ||
                (element.LocalName == "sectionLst" && element.NamespaceURI == SectionNamespace)))
            {
                throw new NotSupportedException("Reordering slides with presentation extensions or sections is not supported.");
            }
            var entry = entries[oldIndex];
            entries.RemoveAt(oldIndex);
            entries.Insert(newIndex, entry);
            slides.RemoveAt(oldIndex);
            slides.Insert(newIndex, slide);
        }

        /// <summary>
        /// Removes a slide at a zero-based position and returns its wrapper.
        /// Dependency parts are retained; the slide part is retained if another relationship targets it.
        /// Unsupported ancillary references are rejected before the document is changed.
        /// Further changes through the removed wrapper are not saved.
        /// </summary>
        public XSLFSlide RemoveSlide(int index)
        {
            EnsureWritable();
            if (index < 0 || index >= slides.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            var entries = GetConsistentSlideEntries();
            var entry = entries[index];
            EnsureCanRemoveSlide(entry);
            var slide = slides[index];
            var part = slide.GetPackagePart();
            var retainPart = HasOtherIncomingRelationships(part, entry.relationshipId);
            if (retainPart)
            {
                // A newly created slide may not have been serialized before it is detached.
                slide.Commit();
            }
            RemoveRelation(slide, false);
            if (!retainPart)
            {
                Package.RemovePart(part.PartName);
                var relationshipsName = PackagingUriHelper.GetRelationshipPartName(part.PartName);
                if (Package.ContainPart(relationshipsName))
                {
                    Package.RemovePart(relationshipsName);
                }
            }
            slide.Committed = true;
            entries.RemoveAt(index);
            slides.RemoveAt(index);
            return slide;
        }

        /// <summary>
        /// Returns the slide dimensions in points (1/72 inch).
        /// </summary>
        public SKSize GetPageSize()
        {
            return new SKSize(
                (float)Units.ToPoints(ReadSlideSize("cx")),
                (float)Units.ToPoints(ReadSlideSize("cy")));
        }

        /// <summary>
        /// Sets custom slide dimensions in points, from 72 through 4032 points per axis.
        /// Shapes are not scaled and notes dimensions are not changed.
        /// </summary>
        public void SetPageSize(SKSize size)
        {
            EnsureWritable();
            var width = ToSlideSize(size.Width, nameof(size));
            var height = ToSlideSize(size.Height, nameof(size));
            var element = presentation.sldSz ?? XSLFXmlHelper.CreateElement(
                "<p:sldSz xmlns:p=\"" + CT_Presentation.Namespace + "\"/>");
            element.SetAttribute("cx", width.ToString(CultureInfo.InvariantCulture));
            element.SetAttribute("cy", height.ToString(CultureInfo.InvariantCulture));
            element.SetAttribute("type", "custom");
            presentation.sldSz = element;
        }

        /// <summary>
        /// Returns embedded objects referenced by internal OLE object or package relationships.
        /// </summary>
        public override List<PackagePart> GetAllEmbedds()
        {
            var embeddings = new List<PackagePart>();
            foreach (var part in Package.GetParts().Where(part => !part.IsRelationshipPart))
            {
                foreach (var relationshipType in new[] { OLE_OBJECT_REL_TYPE, PACK_OBJECT_REL_TYPE })
                {
                    foreach (var relationship in part.GetRelationshipsByType(relationshipType))
                    {
                        if (relationship.TargetMode == TargetMode.Internal)
                        {
                            var embedded = part.GetRelatedPart(relationship);
                            if (!embeddings.Contains(embedded))
                            {
                                embeddings.Add(embedded);
                            }
                        }
                    }
                }
            }
            return embeddings;
        }

        internal override void OnDocumentRead()
        {
            presentation = XSLFXmlHelper.Read<CT_Presentation>(GetPackagePart());
            foreach (var master in GetSlideMasters())
            {
                foreach (var layout in master.GetSlideLayouts())
                {
                    if (!ReferenceEquals(layout.GetSlideMaster(), master))
                    {
                        throw new POIXMLException("Slide layout references a different master: " +
                            layout.GetPackagePart().PartName.Name);
                    }
                }
                master.GetTheme();
            }
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
            GetProperties().ExtendedProperties.GetUnderlyingProperties().Slides = slides.Count;
            XSLFXmlHelper.Write(GetPackagePart(), presentation);
        }

        /// <summary>
        /// Closes the underlying package.
        /// </summary>
        public void Dispose()
        {
            Close();
        }

        private void EnsureWritable()
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
                throw new InvalidOperationException("Cannot modify a read-only presentation.");
            }
        }

        private List<CT_SlideIdListEntry> GetConsistentSlideEntries()
        {
            var entries = presentation.sldIdLst?.sldId;
            if (entries == null || entries.Count != slides.Count)
            {
                throw new POIXMLException("The presentation slide list is inconsistent with the user model.");
            }
            var identifiers = new HashSet<uint>();
            var relationCounts = RelationParts.GroupBy(relation => relation.DocumentPart)
                .ToDictionary(group => group.Key, group => group.Count());
            for (var i = 0; i < entries.Count; i++)
            {
                var entry = entries[i];
                if (entry == null || entry.id < 256 || entry.id > int.MaxValue || !identifiers.Add(entry.id) ||
                    !ReferenceEquals(XSLFXmlHelper.GetReferencedPart<XSLFSlide>(
                        this, entry.relationshipId, XSLFRelation.SLIDE), slides[i]) ||
                    !relationCounts.TryGetValue(slides[i], out var count) || count != 1)
                {
                    throw new POIXMLException("The presentation slide list is inconsistent with the user model.");
                }
            }
            return entries;
        }

        private IEnumerable<XmlElement> GetPresentationElements()
        {
            foreach (var root in presentation.AnyElements ?? Array.Empty<XmlElement>())
            {
                foreach (XmlElement element in root.SelectNodes("descendant-or-self::*"))
                {
                    yield return element;
                }
            }
        }

        private void EnsureCanRemoveSlide(CT_SlideIdListEntry entry)
        {
            if ((presentation.AnyElements ?? Array.Empty<XmlElement>()).Any(root =>
                root.NamespaceURI != CT_Presentation.Namespace))
            {
                throw new NotSupportedException("Removing slides with opaque presentation content is not supported.");
            }
            if ((presentation.AnyAttributes ?? Array.Empty<XmlAttribute>()).Any(attribute =>
                attribute.NamespaceURI == CT_Presentation.RelationshipsNamespace &&
                attribute.Value == entry.relationshipId))
            {
                throw new NotSupportedException("An additional presentation attribute references this slide.");
            }
            foreach (var element in GetPresentationElements())
            {
                if (element.LocalName == "ext" && element.NamespaceURI == CT_Presentation.Namespace &&
                    element.ChildNodes.OfType<XmlElement>().Any(child =>
                        child.LocalName != "sectionLst" || child.NamespaceURI != SectionNamespace))
                {
                    throw new NotSupportedException("Removing slides with opaque presentation extensions is not supported.");
                }
                if ((element.LocalName == "sldId" && element.NamespaceURI == SectionNamespace &&
                    MatchesSlideId(element.GetAttribute("id"), entry.id)) ||
                    element.Attributes.Cast<XmlAttribute>().Any(attribute =>
                        (attribute.NamespaceURI == CT_Presentation.RelationshipsNamespace &&
                            attribute.Value == entry.relationshipId) ||
                        ((attribute.LocalName == "sldId" || attribute.LocalName == "slideId") &&
                            MatchesSlideId(attribute.Value, entry.id))))
                {
                    throw new NotSupportedException("A custom show, section, or presentation extension references this slide.");
                }
            }
        }

        private bool HasOtherIncomingRelationships(PackagePart slide, string relationshipId)
        {
            foreach (var relationship in Package.Relationships)
            {
                if (TargetsPart(relationship, slide.PartName))
                {
                    return true;
                }
            }
            foreach (var source in Package.GetParts().Where(part => !part.IsRelationshipPart && part != slide))
            {
                foreach (var relationship in source.Relationships)
                {
                    if (source == GetPackagePart() && relationship.Id == relationshipId)
                    {
                        continue;
                    }
                    if (TargetsPart(relationship, slide.PartName))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static bool MatchesSlideId(string value, uint slideId)
        {
            return uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) && id == slideId;
        }

        private static bool TargetsPart(PackageRelationship relationship, PackagePartName partName)
        {
            if (relationship.TargetMode != TargetMode.Internal)
            {
                return false;
            }
            var target = relationship.TargetUri.OriginalString;
            var fragment = target.IndexOf('#');
            if (fragment >= 0)
            {
                target = target.Substring(0, fragment);
            }
            return PackagingUriHelper.CreatePartName(target).Equals(partName);
        }

        private int ReadSlideSize(string attribute)
        {
            if (presentation.sldSz == null ||
                !int.TryParse(presentation.sldSz.GetAttribute(attribute), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out var value) || value < MinSlideSize || value > MaxSlideSize)
            {
                throw new POIXMLException("Missing or invalid presentation slide size: " + attribute);
            }
            return value;
        }

        private static int ToSlideSize(float points, string parameterName)
        {
            if (float.IsNaN(points) || float.IsInfinity(points) ||
                points < Units.ToPoints(MinSlideSize) || points > Units.ToPoints(MaxSlideSize))
            {
                throw new ArgumentOutOfRangeException(parameterName,
                    "Slide dimensions must be finite and between 72 and 4032 points.");
            }
            return Units.ToEMU(points);
        }

        private void EnsureCanCreateSlide()
        {
            EnsureWritable();
            if (presentation.AnyElements?.Any(element =>
                element.GetElementsByTagName("sldIdLst", CT_Presentation.Namespace).Count > 0) == true)
            {
                throw new NotSupportedException("Creating slides with opaque wrapped slide lists is not supported.");
            }
        }

        private XSLFSlideLayout FindBlankLayout()
        {
            foreach (var master in GetSlideMasters())
            {
                foreach (var layout in master.GetSlideLayouts())
                {
                    if (layout.Type == "blank")
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

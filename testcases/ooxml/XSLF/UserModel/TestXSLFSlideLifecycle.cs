using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XSLF.UserModel;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFSlideLifecycle
    {
        [TestCase(false, false)]
        [TestCase(false, true)]
        [TestCase(true, false)]
        [TestCase(true, true)]
        public void OpaqueSlideListChildrenBlockLifecycleEditsWithoutMutation(bool alternateContent, bool remove)
        {
            using var presentation = CreatePresentation();
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml(alternateContent
                ? "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" +
                    "<mc:Fallback><p:sldId xmlns:p=\"" + CT_Presentation.Namespace +
                    "\" xmlns:r=\"" + CT_Presentation.RelationshipsNamespace +
                    "\" id=\"258\" r:id=\"" + presentation.GetRelationId(presentation.GetSlides()[2]) +
                    "\"/></mc:Fallback></mc:AlternateContent>"
                : "<future:metadata xmlns:future=\"urn:npoi:test\" slideId=\"258\"/>");
            presentation.GetCTPresentation().sldIdLst.AnyElements = new[] { document.DocumentElement };
            var xml = document.OuterXml;
            if (remove)
            {
                AssertNoMutation<NotSupportedException>(presentation, () => presentation.RemoveSlide(2));
            }
            else
            {
                AssertNoMutation<NotSupportedException>(presentation, () =>
                    presentation.SetSlideOrder(presentation.GetSlides()[2], 0));
            }
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCTPresentation().sldIdLst.AnyElements[0].OuterXml, Is.EqualTo(xml));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(3));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void FormattedGroupedTextSurvivesResizeReorderAndRemoval(bool retainRemovedPart)
        {
            using var presentation = new XMLSlideShow();
            var removed = presentation.CreateSlide();
            var removedBox = removed.CreateTextBox();
            removedBox.SetText("Removed text");
            var removedRun = removedBox.TextParagraphs[0].TextRuns[0];
            removedRun.IsItalic = true;
            var kept = presentation.CreateSlide();
            var group = kept.CreateGroup();
            var groupAnchor = SKRect.Create(10, 20, 300, 200);
            var interiorAnchor = SKRect.Create(0, 0, 600, 400);
            var textAnchor = SKRect.Create(40, 50, 200, 80);
            group.SetAnchor(groupAnchor);
            group.SetInteriorAnchor(interiorAnchor);
            var box = group.CreateTextBox();
            box.SetAnchor(textAnchor);
            box.SetText("Hello \u4e2d\u6587\nsecond");
            var run = box.TextParagraphs[0].TextRuns[0];
            run.IsBold = true;
            run.FontSize = 18;
            run.FontFamily = "+mj-lt";
            var ids = new[] { group.ID, box.ID };
            if (retainRemovedPart)
            {
                kept.GetPackagePart().AddRelationship(removed.GetPackagePart().PartName,
                    TargetMode.Internal, XSLFRelation.SLIDE.Relation, "linkedSlide");
            }

            presentation.SetPageSize(new SKSize(960, 540));
            presentation.SetSlideOrder(kept, 0);
            presentation.RemoveSlide(1);
            Assert.Throws<InvalidOperationException>(() => removedBox.SetText("Must not be saved"));
            Assert.Throws<InvalidOperationException>(() => removedRun.IsItalic = false);
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            Assert.That(again.GetPageSize(), Is.EqualTo(new SKSize(960, 540)));
            Assert.That(again.GetSlides().Count, Is.EqualTo(1));
            Assert.That(again.GetProperties().ExtendedProperties.Slides, Is.EqualTo(1));
            var savedGroup = (XSLFGroupShape)again.GetSlides()[0].GetShapes().Single();
            var savedBox = (XSLFTextBox)savedGroup.GetShapes().Single();
            var savedRun = savedBox.TextParagraphs[0].TextRuns[0];
            Assert.That(new[] { savedGroup.ID, savedBox.ID }, Is.EqualTo(ids));
            Assert.That(savedGroup.GetAnchor(), Is.EqualTo(groupAnchor));
            Assert.That(savedGroup.GetInteriorAnchor(), Is.EqualTo(interiorAnchor));
            Assert.That(savedBox.GetAnchor(), Is.EqualTo(textAnchor));
            Assert.That(savedBox.Text, Is.EqualTo("Hello \u4e2d\u6587\nsecond"));
            Assert.That(savedRun.IsBold, Is.True);
            Assert.That(savedRun.FontSize, Is.EqualTo(18));
            Assert.That(savedRun.FontFamily, Is.EqualTo("Arial"));
            var removedPart = again.Package.GetPart(removed.GetPackagePart().PartName);
            if (retainRemovedPart)
            {
                Assert.That(removedPart, Is.Not.Null);
                var removedXml = ReadXml(removedPart);
                Assert.That(removedXml.GetElementsByTagName("t",
                    "http://schemas.openxmlformats.org/drawingml/2006/main")[0].InnerText,
                    Is.EqualTo("Removed text"));
            }
            else
            {
                Assert.That(removedPart, Is.Null);
            }
            AssertInternalRelationshipsResolve(again.Package);
        }

        [TestCase(0, 2)]
        [TestCase(2, 0)]
        [TestCase(1, 2)]
        [TestCase(1, 0)]
        [TestCase(0, 0)]
        public void ReorderingKeepsSlideIdsRelationshipsAndContents(int oldIndex, int newIndex)
        {
            using var presentation = CreatePresentation();
            var slide = presentation.GetSlides()[oldIndex];
            var entries = presentation.GetCTPresentation().sldIdLst.sldId.ToList();
            var movedEntry = entries[oldIndex];
            entries.RemoveAt(oldIndex);
            entries.Insert(newIndex, movedEntry);
            var names = presentation.GetSlides().Select(item => item.GetXmlObject().cSld.GetAttribute("name")).ToList();
            var movedName = names[oldIndex];
            names.RemoveAt(oldIndex);
            names.Insert(newIndex, movedName);
            var count = presentation.Package.GetParts().Count;

            presentation.SetSlideOrder(slide, newIndex);
            Assert.That(presentation.GetSlides()[newIndex], Is.SameAs(slide));
            Assert.That(presentation.GetCTPresentation().sldIdLst.sldId, Is.EqualTo(entries));
            Assert.That(presentation.Package.GetParts().Count, Is.EqualTo(count));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId.Select(entry => entry.id),
                Is.EqualTo(entries.Select(entry => entry.id)));
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId.Select(entry => entry.relationshipId),
                Is.EqualTo(entries.Select(entry => entry.relationshipId)));
            Assert.That(reopened.GetSlides().Select(item => item.GetXmlObject().cSld.GetAttribute("name")), Is.EqualTo(names));
            Assert.That(reopened.GetProperties().ExtendedProperties.Slides, Is.EqualTo(3));
        }

        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        public void RemovalDeletesOnlyTheUnreferencedSlideAndItsRelationships(int index)
        {
            using var original = CreatePresentation();
            using var presentation = Open(Save(original));
            var slide = presentation.GetSlides()[index];
            var name = slide.GetPackagePart().PartName;
            var relationshipId = presentation.GetRelationId(slide);
            var expected = presentation.GetCTPresentation().sldIdLst.sldId.Where((entry, i) => i != index).ToArray();
            var master = presentation.GetSlideMasters()[0];
            var layout = slide.GetSlideLayout();
            var theme = slide.GetTheme();
            Assert.That(presentation.RemoveSlide(index), Is.SameAs(slide));
            Assert.That(presentation.Package.ContainPart(name), Is.False);
            Assert.That(presentation.Package.ContainPart(PackagingUriHelper.GetRelationshipPartName(name)), Is.False);
            Assert.That(presentation.GetPackagePart().GetRelationship(relationshipId), Is.Null);
            Assert.That(presentation.GetRelationById(relationshipId), Is.Null);
            Assert.That(presentation.Package.ContainPart(master.GetPackagePart().PartName), Is.True);
            Assert.That(presentation.Package.ContainPart(layout.GetPackagePart().PartName), Is.True);
            Assert.That(presentation.Package.ContainPart(theme.GetPackagePart().PartName), Is.True);

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(2));
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId.Select(entry => entry.id),
                Is.EqualTo(expected.Select(entry => entry.id)));
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId.Select(entry => entry.relationshipId),
                Is.EqualTo(expected.Select(entry => entry.relationshipId)));
            Assert.That(reopened.GetProperties().ExtendedProperties.Slides, Is.EqualTo(2));
            AssertInternalRelationshipsResolve(reopened.Package);
        }

        [Test]
        public void AllSlidesCanBeRemovedAndNewSlidesCanThenBeCreated()
        {
            using var presentation = CreatePresentation();
            var removed = presentation.RemoveSlide(0);
            AssertNoMutation<ArgumentException>(presentation, () => presentation.SetSlideOrder(removed, 0));
            while (presentation.GetSlides().Count > 0)
            {
                presentation.RemoveSlide(0);
            }
            using var empty = Open(Save(presentation));
            Assert.That(empty.GetSlides(), Is.Empty);
            Assert.That(empty.GetProperties().ExtendedProperties.Slides, Is.EqualTo(0));
            empty.CreateSlide();
            using var reopened = Open(Save(empty));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(1));
            Assert.That(reopened.GetProperties().ExtendedProperties.Slides, Is.EqualTo(1));
            AssertInternalRelationshipsResolve(reopened.Package);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void IncomingRelationshipsRetainTheRemovedSlidePart(bool fragment)
        {
            using var presentation = CreatePresentation();
            var removed = presentation.GetSlides()[0];
            var source = presentation.GetSlides()[1].GetPackagePart();
            var name = removed.GetPackagePart().PartName;
            source.AddRelationship(new Uri(name.Name + (fragment ? "#bookmark" : ""), UriKind.Relative),
                TargetMode.Internal, CT_Presentation.RelationshipsNamespace + "/slide", "linkedSlide");
            presentation.RemoveSlide(0);
            Assert.That(presentation.Package.ContainPart(name), Is.True);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(2));
            var part = reopened.Package.GetPart(name);
            Assert.That(part, Is.Not.Null);
            var linked = reopened.GetSlides()[0].GetPackagePart().GetRelationship("linkedSlide");
            Assert.That(linked.TargetUri.OriginalString, Is.EqualTo(name.Name + (fragment ? "#bookmark" : "")));
            Assert.That(ReadXml(part).DocumentElement["cSld", CT_Presentation.Namespace].GetAttribute("name"), Is.EqualTo("0"));
            AssertInternalRelationshipsResolve(reopened.Package);
        }

        [Test]
        public void NotesBackReferencesRetainAnUnwrittenSlideWithoutKeepingItInTheSlideList()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            slide.GetXmlObject().cSld.SetAttribute("name", "Retained for notes");
            var name = slide.GetPackagePart().PartName;
            var notes = presentation.Package.CreatePart(PackagingUriHelper.CreatePartName("/ppt/notesSlides/notesSlide1.xml"),
                "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml");
            WriteBytes(notes, Encoding.UTF8.GetBytes("<p:notes xmlns:p=\"" + CT_Presentation.Namespace + "\">" +
                slide.GetXmlObject().cSld.OuterXml + "</p:notes>"));
            var notesBytes = ReadBytes(notes);
            notes.AddRelationship(name, TargetMode.Internal, XSLFRelation.SLIDE.Relation);
            slide.GetPackagePart().AddRelationship(notes.PartName, TargetMode.Internal,
                CT_Presentation.RelationshipsNamespace + "/notesSlide");

            presentation.RemoveSlide(0);
            slide.GetXmlObject().cSld.SetAttribute("name", "Changes to a removed wrapper are not saved");
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides(), Is.Empty);
            Assert.That(ReadBytes(reopened.Package.GetPart(notes.PartName)), Is.EqualTo(notesBytes));
            Assert.That(ReadXml(reopened.Package.GetPart(name)).DocumentElement["cSld", CT_Presentation.Namespace]
                .GetAttribute("name"), Is.EqualTo("Retained for notes"));
            AssertInternalRelationshipsResolve(reopened.Package);
        }

        [Test]
        public void SharedMediaAndUnreferencedDependenciesAreNotRecursivelyDeleted()
        {
            using var original = CreatePresentation();
            var media = original.Package.CreatePart(PackagingUriHelper.CreatePartName("/ppt/media/shared.bin"),
                "application/octet-stream");
            var bytes = new byte[] { 0, 1, 2, 127, 255 };
            WriteBytes(media, bytes);
            foreach (var slide in original.GetSlides())
            {
                slide.GetPackagePart().AddRelationship(media.PartName, TargetMode.Internal,
                    CT_Presentation.RelationshipsNamespace + "/image", "sharedMedia");
            }
            using var presentation = Open(Save(original));
            presentation.RemoveSlide(0);
            using var reopened = Open(Save(presentation));
            Assert.That(ReadBytes(reopened.Package.GetPart(media.PartName)), Is.EqualTo(bytes));
            foreach (var slide in reopened.GetSlides())
            {
                Assert.That(slide.GetPackagePart().GetRelationship("sharedMedia"), Is.Not.Null);
            }
            reopened.RemoveSlide(0);
            reopened.RemoveSlide(0);
            using var empty = Open(Save(reopened));
            Assert.That(ReadBytes(empty.Package.GetPart(media.PartName)), Is.EqualTo(bytes));
            Assert.That(empty.GetSlideMasters().Count, Is.EqualTo(1));
        }

        [Test]
        public void RealPresentationCanBeReorderedAndReducedWithoutChangingDependencyBytes()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("sample.pptx");
            using var presentation = new XMLSlideShow(input);
            var dependencies = presentation.Package.GetParts().Where(part =>
                !part.IsRelationshipPart &&
                part.ContentType != XSLFRelation.MAIN.ContentType &&
                part.ContentType != XSLFRelation.SLIDE.ContentType &&
                !part.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(part => part.PartName.Name, ReadBytes);
            var last = presentation.GetSlides().Last();
            var id = presentation.GetRelationId(last);
            presentation.SetSlideOrder(last, 0);
            presentation.RemoveSlide(1);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId[0].relationshipId, Is.EqualTo(id));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(1));
            foreach (var dependency in dependencies)
            {
                Assert.That(ReadBytes(reopened.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key))),
                    Is.EqualTo(dependency.Value), dependency.Key);
            }
            AssertInternalRelationshipsResolve(reopened.Package);
        }

        [Test]
        public void AmbiguousCachedRelationshipsAreRejectedBeforeRemovingASlide()
        {
            using var original = CreatePresentation();
            original.GetPackagePart().AddRelationship(original.GetSlides()[0].GetPackagePart().PartName,
                TargetMode.Internal, XSLFRelation.SLIDE.Relation, "duplicateSlide");
            using var presentation = Open(Save(original));
            AssertNoMutation<POIXMLException>(presentation, () => presentation.RemoveSlide(0));
        }

        [TestCase(-1)]
        [TestCase(3)]
        public void InvalidIndexesDoNotChangeTheDocument(int index)
        {
            using var presentation = CreatePresentation();
            AssertNoMutation<ArgumentOutOfRangeException>(presentation, () => presentation.RemoveSlide(index));
            AssertNoMutation<ArgumentOutOfRangeException>(presentation, () =>
                presentation.SetSlideOrder(presentation.GetSlides()[0], index));
        }

        [Test]
        public void NullAndForeignSlidesAreRejectedWithoutMutation()
        {
            using var presentation = CreatePresentation();
            using var other = CreatePresentation();
            AssertNoMutation<ArgumentNullException>(presentation, () => presentation.SetSlideOrder(null, 0));
            AssertNoMutation<ArgumentException>(presentation, () => presentation.SetSlideOrder(other.GetSlides()[0], 0));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void InconsistentLowLevelSlideListsFailBeforeMutation(bool reverse)
        {
            using var presentation = CreatePresentation();
            var entries = presentation.GetCTPresentation().sldIdLst.sldId;
            if (reverse)
            {
                entries.Reverse();
            }
            else
            {
                entries.RemoveAt(0);
            }
            AssertNoMutation<POIXMLException>(presentation, () => presentation.RemoveSlide(0));
            AssertNoMutation<POIXMLException>(presentation, () =>
                presentation.SetSlideOrder(presentation.GetSlides()[0], 1));
        }

        [Test]
        public void ReadOnlyAndClosedPresentationsRejectLifecycleChanges()
        {
            using var original = CreatePresentation();
            using var package = OPCPackage.Open(new MemoryStream(Save(original)), true);
            using var readOnly = new XMLSlideShow(package);
            AssertNoMutation<InvalidOperationException>(readOnly, () => readOnly.RemoveSlide(0));
            AssertNoMutation<InvalidOperationException>(readOnly, () => readOnly.SetSlideOrder(readOnly.GetSlides()[0], 1));
            var slide = original.GetSlides()[0];
            original.Dispose();
            Assert.Throws<ObjectDisposedException>(() => original.RemoveSlide(0));
            Assert.Throws<ObjectDisposedException>(() => original.SetSlideOrder(slide, 1));
        }

        [Test]
        public void ACustomShowBlocksRemovalOfItsReferencedSlideButNotOtherSlidesOrReordering()
        {
            using var presentation = CreatePresentation();
            var slide = presentation.GetSlides()[0];
            var id = presentation.GetRelationId(slide);
            AddPresentationElement(presentation,
                "<p:custShowLst xmlns:p=\"" + CT_Presentation.Namespace + "\" xmlns:r=\"" +
                CT_Presentation.RelationshipsNamespace + "\"><p:custShow name=\"Show\" id=\"1\">" +
                "<p:sldLst><p:sld r:id=\"" + id + "\"/></p:sldLst></p:custShow></p:custShowLst>");
            AssertNoMutation<NotSupportedException>(presentation, () => presentation.RemoveSlide(0));
            presentation.RemoveSlide(2);
            presentation.SetSlideOrder(slide, 1);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(2));
            Assert.That(reopened.GetPackagePart().GetRelationship(id), Is.Not.Null);
        }

        [TestCase("256")]
        [TestCase("0256")]
        public void SectionReferencesBlockRemovalAndSectionedPresentationsCannotBeReordered(string slideId)
        {
            using var presentation = CreatePresentation();
            AddPresentationElement(presentation,
                "<p:extLst xmlns:p=\"" + CT_Presentation.Namespace + "\" xmlns:p14=\"" +
                "http://schemas.microsoft.com/office/powerpoint/2010/main\"><p:ext uri=\"test\">" +
                "<p14:sectionLst><p14:section name=\"Section\" id=\"{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\">" +
                "<p14:sldIdLst><p14:sldId id=\"" + slideId + "\"/></p14:sldIdLst>" +
                "</p14:section></p14:sectionLst></p:ext></p:extLst>");
            AssertNoMutation<NotSupportedException>(presentation, () => presentation.RemoveSlide(0));
            AssertNoMutation<NotSupportedException>(presentation, () =>
                presentation.SetSlideOrder(presentation.GetSlides()[0], 1));
            presentation.RemoveSlide(2);
            Assert.That(presentation.GetSlides().Count, Is.EqualTo(2));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void OpaquePresentationContentIsRejectedConservatively(bool alternateContent)
        {
            using var presentation = CreatePresentation();
            AddPresentationElement(presentation, alternateContent
                ? "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"/>"
                : "<p:extLst xmlns:p=\"" + CT_Presentation.Namespace +
                    "\" xmlns:x=\"urn:npoi:test\"><p:ext uri=\"test\"><x:unknown/></p:ext></p:extLst>");
            AssertNoMutation<NotSupportedException>(presentation, () => presentation.RemoveSlide(0));
            AssertNoMutation<NotSupportedException>(presentation, () =>
                presentation.SetSlideOrder(presentation.GetSlides()[0], 1));
        }

        private static XMLSlideShow CreatePresentation()
        {
            var presentation = new XMLSlideShow();
            for (var i = 0; i < 3; i++)
            {
                presentation.CreateSlide().GetXmlObject().cSld.SetAttribute("name", i.ToString());
            }
            return presentation;
        }

        private static void AssertNoMutation<T>(XMLSlideShow presentation, Action action) where T : Exception
        {
            var parts = presentation.Package.GetParts().Count;
            var relationships = presentation.GetPackagePart().Relationships.Size;
            var entries = presentation.GetCTPresentation().sldIdLst.sldId.ToArray();
            var slides = presentation.GetSlides().ToArray();
            Assert.Throws<T>(action);
            Assert.That(presentation.Package.GetParts().Count, Is.EqualTo(parts));
            Assert.That(presentation.GetPackagePart().Relationships.Size, Is.EqualTo(relationships));
            Assert.That(presentation.GetCTPresentation().sldIdLst.sldId, Is.EqualTo(entries));
            Assert.That(presentation.GetSlides(), Is.EqualTo(slides));
        }

        private static void AddPresentationElement(XMLSlideShow presentation, string xml)
        {
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml(xml);
            presentation.GetCTPresentation().AnyElements =
                (presentation.GetCTPresentation().AnyElements ?? Array.Empty<XmlElement>())
                .Concat(new[] { document.DocumentElement }).ToArray();
        }

        private static void AssertInternalRelationshipsResolve(OPCPackage package)
        {
            foreach (var part in package.GetParts().Where(part => !part.IsRelationshipPart))
            {
                foreach (var relationship in part.Relationships)
                {
                    if (relationship.TargetMode == TargetMode.Internal)
                    {
                        var target = relationship.TargetUri.OriginalString.Split('#')[0];
                        Assert.That(package.GetPart(PackagingUriHelper.CreatePartName(target)), Is.Not.Null);
                    }
                }
            }
        }

        private static XMLSlideShow Open(byte[] bytes)
        {
            return new XMLSlideShow(new MemoryStream(bytes));
        }

        private static byte[] Save(XMLSlideShow presentation)
        {
            using var output = new MemoryStream();
            presentation.Write(output);
            return output.ToArray();
        }

        private static byte[] ReadBytes(PackagePart part)
        {
            using var input = part.GetInputStream();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        private static void WriteBytes(PackagePart part, byte[] bytes)
        {
            using var output = part.GetOutputStream();
            output.Write(bytes, 0, bytes.Length);
        }

        private static XmlDocument ReadXml(PackagePart part)
        {
            using var stream = part.GetInputStream();
            var document = new XmlDocument { XmlResolver = null };
            document.Load(stream);
            return document;
        }
    }
}

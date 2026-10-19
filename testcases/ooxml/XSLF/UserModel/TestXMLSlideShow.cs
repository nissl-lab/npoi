using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XSLF.UserModel;
using NPOI.XWPF.UserModel;
using NUnit.Framework;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXMLSlideShow
    {
        [Test]
        public void EmptyPresentationCanBeSavedAndReopened()
        {
            using var presentation = new XMLSlideShow();
            Assert.That(presentation.GetSlides(), Is.Empty);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides(), Is.Empty);
            Assert.That(reopened.GetPackagePart().ContentType, Is.EqualTo(XSLFRelation.MAIN.ContentType));
            Assert.That(reopened.GetCTPresentation().sldSz.GetAttribute("cx"), Is.EqualTo("9144000"));
            Assert.That(reopened.GetCTPresentation().notesSz, Is.Not.Null);
        }

        [Test]
        public void CreatedSlidesHaveCompleteLayoutMasterAndThemeRelationships()
        {
            using var presentation = new XMLSlideShow();
            var first = presentation.CreateSlide();
            var second = presentation.CreateSlide();
            Assert.That(presentation.GetSlides(), Is.EqualTo(new[] { first, second }));
            Assert.That(((ICollection<XSLFSlide>)presentation.GetSlides()).IsReadOnly, Is.True);
            Assert.That(presentation.GetCTPresentation().sldIdLst.sldId.Select(entry => entry.id),
                Is.EqualTo(new uint[] { 256, 257 }));

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(2));
            foreach (var slide in reopened.GetSlides())
            {
                var layout = RelatedPart(slide.GetPackagePart(), XSLFRelation.SLIDE_LAYOUT);
                var master = RelatedPart(layout, XSLFRelation.SLIDE_MASTER);
                var theme = RelatedPart(master, XSLFRelation.THEME);
                Assert.That(RelatedPart(master, XSLFRelation.SLIDE_LAYOUT), Is.SameAs(layout));
                Assert.That(RelatedPart(reopened.GetPackagePart(), XSLFRelation.SLIDE_MASTER), Is.SameAs(master));
                Assert.That(theme.ContentType, Is.EqualTo(XSLFRelation.THEME.ContentType));
                Assert.That(ReadXml(layout).DocumentElement.GetAttribute("type"), Is.EqualTo("blank"));
                Assert.That(slide.GetXmlObject().cSld["spTree", CT_Presentation.Namespace], Is.Not.Null);
            }
        }

        [Test]
        public void SlideOrderComesFromThePresentationRatherThanPartNames()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide().GetXmlObject().cSld.SetAttribute("name", "First");
            presentation.CreateSlide().GetXmlObject().cSld.SetAttribute("name", "Second");
            presentation.GetCTPresentation().sldIdLst.sldId.Reverse();

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Select(slide => slide.GetXmlObject().cSld.GetAttribute("name")),
                Is.EqualTo(new[] { "Second", "First" }));
            reopened.CreateSlide();
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId.Select(entry => entry.id),
                Is.EqualTo(new uint[] { 257, 256, 258 }));
        }

        [Test]
        public void RepeatedSavesAndAppendingAfterReopeningDoNotTruncateParts()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            var first = Save(presentation);
            var second = Save(presentation);
            using var reopened = Open(second);
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(1));
            reopened.CreateSlide();
            using var again = Open(Save(reopened));
            Assert.That(again.GetSlides().Count, Is.EqualTo(2));
            using var original = Open(first);
            Assert.That(original.GetSlides().Count, Is.EqualTo(1));
            Assert.That(ReadBytes(RelatedPart(again.GetPackagePart(), XSLFRelation.SLIDE_MASTER)),
                Is.EqualTo(ReadBytes(RelatedPart(original.GetPackagePart(), XSLFRelation.SLIDE_MASTER))));
        }

        [Test]
        public void AnExistingPresentationCanBeExtendedUsingItsBlankLayout()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("sample.pptx");
            using var presentation = new XMLSlideShow(input);
            var count = presentation.GetSlides().Count;
            var originalIds = presentation.GetCTPresentation().sldIdLst.sldId.Select(entry => entry.id).ToArray();
            var slide = presentation.CreateSlide();
            var layout = RelatedPart(slide.GetPackagePart(), XSLFRelation.SLIDE_LAYOUT);
            Assert.That(ReadXml(layout).DocumentElement.GetAttribute("type"), Is.EqualTo("blank"));

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(count + 1));
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId.Take(count).Select(entry => entry.id),
                Is.EqualTo(originalIds));
        }

        [Test]
        public void NewSlidePartNamesDoNotCollideWithExistingParts()
        {
            using var presentation = new XMLSlideShow();
            var occupiedName = PackagingUriHelper.CreatePartName("/ppt/slides/slide1.xml");
            presentation.Package.CreatePart(occupiedName, "application/octet-stream");
            var slide = presentation.CreateSlide();
            Assert.That(slide.GetPackagePart().PartName.Name, Is.EqualTo("/ppt/slides/slide2.xml"));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(1));
            Assert.That(reopened.Package.GetPart(occupiedName).ContentType, Is.EqualTo("application/octet-stream"));
        }

        [TestCase("sample.pptx")]
        [TestCase("shapes.pptx")]
        [TestCase("WithMaster.pptx")]
        [TestCase("pie-chart.pptx")]
        [TestCase("with_japanese.pptx")]
        [TestCase("layouts.pptx")]
        [TestCase("backgrounds.pptx")]
        [TestCase("themes.pptx")]
        public void ExistingPresentationRetainsUnmodeledXmlAndPackageParts(string name)
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(name);
            using var presentation = new XMLSlideShow(stream);
            var slideXml = presentation.GetSlides().Select(slide => ReadXml(slide.GetPackagePart()).OuterXml).ToArray();
            var originalPresentation = ReadXml(presentation.GetPackagePart()).OuterXml;
            var parts = presentation.Package.GetParts().Where(part =>
                !part.IsRelationshipPart && part.ContentType != XSLFRelation.MAIN.ContentType &&
                part.ContentType != XSLFRelation.SLIDE.ContentType &&
                !part.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(part => part.PartName.Name, ReadBytes);

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Count, Is.EqualTo(slideXml.Length));
            AssertXmlEquivalent(originalPresentation, ReadXml(reopened.GetPackagePart()).OuterXml);
            for (var i = 0; i < slideXml.Length; i++)
            {
                AssertXmlEquivalent(slideXml[i], ReadXml(reopened.GetSlides()[i].GetPackagePart()).OuterXml);
            }
            foreach (var part in parts)
            {
                var savedPart = reopened.Package.GetPart(PackagingUriHelper.CreatePartName(part.Key));
                Assert.That(savedPart, Is.Not.Null, part.Key);
                Assert.That(ReadBytes(savedPart), Is.EqualTo(part.Value), part.Key);
            }
        }

        [Test]
        public void UnknownPartsAndExternalRelationshipsSurviveMultipleSaves()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var embedded = presentation.Package.CreatePart(
                PackagingUriHelper.CreatePartName("/ppt/embeddings/example.bin"), "application/octet-stream");
            var payload = new byte[] { 0, 1, 127, 128, 255 };
            using (var output = embedded.GetOutputStream())
            {
                output.Write(payload, 0, payload.Length);
            }
            slide.GetPackagePart().AddRelationship(embedded.PartName, TargetMode.Internal,
                CT_Presentation.RelationshipsNamespace + "/oleObject");
            slide.GetPackagePart().AddExternalRelationship("https://example.com/", XSLFRelation.THEME.Relation, "external");

            using var reopened = Open(Save(presentation));
            Save(reopened);
            using var again = Open(Save(reopened));
            Assert.That(again.GetAllEmbedds().Count, Is.EqualTo(1));
            Assert.That(ReadBytes(again.GetAllEmbedds()[0]), Is.EqualTo(payload));
            var external = again.GetSlides()[0].GetPackagePart().GetRelationship("external");
            Assert.That(external.TargetMode, Is.EqualTo(TargetMode.External));
            Assert.That(external.TargetUri.ToString(), Is.EqualTo("https://example.com/"));
        }

        [Test]
        public void MissingSlideRelationshipFailsExplicitly()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            presentation.GetCTPresentation().sldIdLst.sldId[0].relationshipId = "missing";
            AssertInvalid(Save(presentation), "Invalid slide relationship");
        }

        [TestCase(false)]
        [TestCase(true)]
        public void IncorrectSlideRelationshipTypesAndExternalSlideTargetsAreRejected(bool external)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var id = presentation.GetRelationId(slide);
            var main = presentation.GetPackagePart();
            main.RemoveRelationship(id);
            if (external)
            {
                main.AddExternalRelationship("https://example.com/slide.xml", XSLFRelation.SLIDE.Relation, id);
            }
            else
            {
                main.AddRelationship(slide.GetPackagePart().PartName, TargetMode.Internal,
                    XSLFRelation.SLIDE_LAYOUT.Relation, id);
            }
            AssertInvalid(Save(presentation), "Invalid slide relationship");
        }

        [Test]
        public void MissingSlidePartFailsExplicitly()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var bytes = Save(presentation);
            using var package = OPCPackage.Open(new MemoryStream(bytes));
            package.RemovePart(slide.GetPackagePart().PartName);
            var error = Assert.Throws<POIXMLException>(() => new XMLSlideShow(package));
            Assert.That(error.Message, Does.Contain("Invalid slide relationship"));
        }

        [TestCase(0U)]
        [TestCase(255U)]
        [TestCase(2147483648U)]
        public void OutOfRangeSlideIdentifiersFailExplicitly(uint id)
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            presentation.GetCTPresentation().sldIdLst.sldId[0].id = id;
            AssertInvalid(Save(presentation), "Invalid or duplicate");
        }

        [Test]
        public void DuplicateSlideIdentifiersFailExplicitly()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            presentation.CreateSlide();
            presentation.GetCTPresentation().sldIdLst.sldId[1].id = 256;
            AssertInvalid(Save(presentation), "Invalid or duplicate");
        }

        [Test]
        public void ASlideCannotBeListedTwice()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            presentation.CreateSlide();
            var entries = presentation.GetCTPresentation().sldIdLst.sldId;
            entries[1].relationshipId = entries[0].relationshipId;
            AssertInvalid(Save(presentation), "referenced more than once");
        }

        [Test]
        public void SlideIdentifierExhaustionDoesNotChangeThePackage()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            presentation.GetCTPresentation().sldIdLst.sldId[0].id = int.MaxValue;
            var count = presentation.Package.GetParts().Count;
            Assert.Throws<InvalidOperationException>(() => presentation.CreateSlide());
            Assert.That(presentation.Package.GetParts().Count, Is.EqualTo(count));
            Assert.That(presentation.GetSlides().Count, Is.EqualTo(1));
        }

        [Test]
        public void MissingCommonSlideDataFailsExplicitly()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide().GetXmlObject().cSld = null;
            AssertInvalid(Save(presentation), "no common slide data");
        }

        [Test]
        public void MissingBlankLayoutFailsWithoutAddingASlide()
        {
            using var presentation = new XMLSlideShow();
            var layout = RelatedPart(RelatedPart(presentation.GetPackagePart(), XSLFRelation.SLIDE_MASTER),
                XSLFRelation.SLIDE_LAYOUT);
            var document = ReadXml(layout);
            document.DocumentElement.SetAttribute("type", "title");
            layout.Clear();
            using (var output = layout.GetOutputStream())
            {
                document.Save(output);
            }
            var count = presentation.Package.GetParts().Count;
            Assert.Throws<NotSupportedException>(() => presentation.CreateSlide());
            Assert.That(presentation.GetSlides(), Is.Empty);
            Assert.That(presentation.Package.GetParts().Count, Is.EqualTo(count));
        }

        [Test]
        public void OtherOfficeDocumentTypesAreRejected()
        {
            using var word = new XWPFDocument();
            using var output = new MemoryStream();
            word.Write(output);
            using var package = OPCPackage.Open(new MemoryStream(output.ToArray()));
            Assert.Throws<POIXMLException>(() => new XMLSlideShow(package));
        }

        [Test]
        public void ReadOnlyPackagesCannotCreateSlides()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            using var package = OPCPackage.Open(new MemoryStream(Save(presentation)), true);
            using var readOnly = new XMLSlideShow(package);
            Assert.That(readOnly.GetSlides().Count, Is.EqualTo(1));
            Assert.Throws<InvalidOperationException>(() => readOnly.CreateSlide());
        }

        [Test]
        public void ClosedPresentationsCannotCreateSlidesOrWrite()
        {
            var presentation = new XMLSlideShow();
            presentation.Dispose();
            presentation.Dispose();
            Assert.Throws<ObjectDisposedException>(() => presentation.CreateSlide());
            using var output = new MemoryStream();
            Assert.Throws<IOException>(() => presentation.Write(output));
        }

        [Test]
        public void DocumentTypeDeclarationsAreRejected()
        {
            using var presentation = new XMLSlideShow();
            using var package = OPCPackage.Open(new MemoryStream(Save(presentation)));
            var main = package.GetPart(PackagingUriHelper.CreatePartName(XSLFRelation.MAIN.DefaultFileName));
            main.Clear();
            using (var output = main.GetOutputStream())
            {
                var xml = Encoding.UTF8.GetBytes(
                    "<!DOCTYPE presentation [<!ENTITY example 'example'>]>" +
                    "<p:presentation xmlns:p=\"" + CT_Presentation.Namespace + "\"/>");
                output.Write(xml, 0, xml.Length);
            }
            Assert.Throws<POIXMLException>(() => new XMLSlideShow(package));
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

        private static void AssertInvalid(byte[] bytes, string message)
        {
            using var package = OPCPackage.Open(new MemoryStream(bytes));
            var error = Assert.Throws<POIXMLException>(() => new XMLSlideShow(package));
            Assert.That(error.Message, Does.Contain(message));
        }

        private static PackagePart RelatedPart(PackagePart source, XSLFRelation descriptor)
        {
            var relationships = source.GetRelationshipsByType(descriptor.Relation);
            Assert.That(relationships.Size, Is.EqualTo(1));
            var relationship = relationships.GetRelationship(0);
            Assert.That(relationship.TargetMode, Is.EqualTo(TargetMode.Internal));
            var target = source.GetRelatedPart(relationship);
            Assert.That(target.ContentType, Is.EqualTo(descriptor.ContentType));
            return target;
        }

        private static XmlDocument ReadXml(PackagePart part)
        {
            using var input = part.GetInputStream();
            var document = new XmlDocument { XmlResolver = null };
            document.Load(input);
            return document;
        }

        private static byte[] ReadBytes(PackagePart part)
        {
            using var input = part.GetInputStream();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        private static void AssertXmlEquivalent(string expected, string actual)
        {
            var first = XElement.Parse(expected);
            var second = XElement.Parse(actual);
            first.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            second.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            Assert.That(XNode.DeepEquals(first, second), Is.True, "PresentationML content changed.");
        }
    }
}

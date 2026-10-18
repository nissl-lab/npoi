using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.SS.UserModel;
using NPOI.XSLF.UserModel;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFPictures
    {
        private const string P = CT_Presentation.Namespace;
        private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string R = CT_Presentation.RelationshipsNamespace;

        [TestCase(PictureType.PNG, "image/png", "png")]
        [TestCase(PictureType.JPEG, "image/jpeg", "jpeg")]
        public void ImagesKeepOriginalBytesMetadataAndPixelsAcrossRepeatedSaves(
            PictureType type, string mime, string extension)
        {
            using var presentation = new XMLSlideShow();
            var data = Image(type);
            var picture = presentation.AddPicture(data, type);
            Assert.That(picture, Is.InstanceOf<IPictureData>());
            Assert.That(picture.Data, Is.EqualTo(data));
            Assert.That(picture.MimeType, Is.EqualTo(mime));
            Assert.That(picture.PictureType, Is.EqualTo(type));
            Assert.That(picture.SuggestFileExtension(), Is.EqualTo(extension));
            Assert.That(picture.GetImageDimensionInPixels(), Is.EqualTo(new SKSizeI(3, 2)));
            presentation.CreateSlide().CreatePicture(picture).SetAnchor(SKRect.Create(-5, 12, 90, 60));
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            var saved = (XSLFPictureShape)again.GetSlides()[0].GetShapes().Single();
            Assert.That(saved.GetAnchor(), Is.EqualTo(SKRect.Create(-5, 12, 90, 60)));
            Assert.That(saved.GetPictureData(), Is.SameAs(again.GetPictureData().Single()));
            Assert.That(saved.GetPictureData().Data, Is.EqualTo(data));
            Assert.That(saved.GetPictureData().GetImageDimensionInPixels(), Is.EqualTo(new SKSizeI(3, 2)));
        }

        [TestCase("slide")]
        [TestCase("layout")]
        [TestCase("master")]
        [TestCase("group")]
        [TestCase("nested")]
        public void AllContainersCreateTypedPicturesWithUniqueIdsAndEmptyAnchors(string kind)
        {
            using var presentation = new XMLSlideShow();
            var container = Container(presentation, kind);
            var data = presentation.AddPicture(Image(PictureType.PNG), PictureType.PNG);
            var snapshot = container.GetShapes();
            var first = container.CreatePicture(data);
            var second = container.CreatePicture(data);
            Assert.That(first.Parent, Is.SameAs(container));
            Assert.That(first.GetAnchor(), Is.EqualTo(SKRect.Empty));
            Assert.That(first.ID, Is.Not.EqualTo(second.ID));
            Assert.That(first.GetPictureData(), Is.SameAs(data));
            Assert.That(snapshot, Is.Empty);
            Assert.That(((ICollection<XSLFShape>)container.GetShapes()).IsReadOnly, Is.True);
            var sheet = first.GetSheet();
            Assert.That(sheet.GetPackagePart().GetRelationshipsByType(PackageRelationshipTypes.IMAGE_PART).Size,
                Is.EqualTo(1));
            var ids = AllShapes(sheet).Select(shape => shape.ID).ToArray();
            Assert.That(ids.Distinct().Count(), Is.EqualTo(ids.Length));
            first.Name = "Picture <&>";
            first.SetAnchor(SKRect.Create(10, 20, 30, 40));
            using var reopened = Open(Save(presentation));
            var saved = Container(reopened, kind).GetShapes().Cast<XSLFPictureShape>().ToArray();
            Assert.That(saved.Select(shape => shape.ID), Is.EqualTo(new[] { first.ID, second.ID }));
            Assert.That(saved[0].Name, Is.EqualTo("Picture <&>"));
            Assert.That(saved[0].GetAnchor(), Is.EqualTo(SKRect.Create(10, 20, 30, 40)));
            Assert.That(saved[0].IsExternalLinkedPicture, Is.False);
            Assert.That(saved[0].GetPictureLink(), Is.Null);
        }

        [Test]
        public void ExactBytesShareMediaAcrossSheetsAndGroupsButDoNotShareShapeIds()
        {
            using var presentation = new XMLSlideShow();
            var bytes = Image(PictureType.PNG);
            var snapshot = presentation.GetPictureData();
            var first = presentation.AddPicture(bytes, PictureType.PNG);
            var second = presentation.AddPicture((byte[])bytes.Clone(), PictureType.PNG);
            Assert.That(second, Is.SameAs(first));
            Assert.That(snapshot, Is.Empty);
            Assert.That(((ICollection<XSLFPictureData>)presentation.GetPictureData()).IsReadOnly, Is.True);
            var slide = presentation.CreateSlide();
            slide.CreatePicture(first);
            slide.CreateGroup().CreateGroup().CreatePicture(second);
            presentation.CreateSlide().CreatePicture(first);
            var master = presentation.GetSlideMasters()[0];
            master.CreatePicture(first);
            master.GetSlideLayouts()[0].CreatePicture(first);
            Assert.That(presentation.GetPictureData().Count, Is.EqualTo(1));
            using var reopened = Open(Save(presentation));
            var images = reopened.GetPictureData();
            Assert.That(images.Count, Is.EqualTo(1));
            Assert.That(reopened.AddPicture(bytes, PictureType.PNG), Is.SameAs(images[0]));
            Assert.That(AllShapes(reopened.GetSlides()[0]).OfType<XSLFPictureShape>()
                .Select(shape => shape.GetPictureData()), Is.All.SameAs(images[0]));
            var removed = reopened.RemoveSlide(0);
            Assert.Throws<InvalidOperationException>(() => removed.CreatePicture(images[0]));
            using var again = Open(Save(reopened));
            Assert.That(again.GetPictureData().Single().Data, Is.EqualTo(bytes));
            Assert.That(((XSLFPictureShape)again.GetSlides()[0].GetShapes().Single()).GetPictureData().Data,
                Is.EqualTo(bytes));
        }

        [Test]
        public void DistinctBytesAndMimeTypesDoNotDeduplicateAndOccupiedNamesAreNotOverwritten()
        {
            using var presentation = new XMLSlideShow();
            var bytes = Image(PictureType.PNG);
            var occupied = Part(presentation, "/ppt/media/image1.png", "application/octet-stream", bytes);
            var wrongMime = Part(presentation, "/ppt/media/wrong.jpeg", "image/jpeg", bytes);
            var picture = presentation.AddPicture(bytes, PictureType.PNG);
            Assert.That(picture.GetPackagePart().PartName.Name, Is.EqualTo("/ppt/media/image2.png"));
            Assert.That(presentation.AddPicture(Image(PictureType.PNG, SKColors.Red), PictureType.PNG),
                Is.Not.SameAs(picture));
            Assert.That(presentation.AddPicture(bytes, PictureType.PNG), Is.SameAs(picture));
            Assert.That(Read(occupied), Is.EqualTo(bytes));
            Assert.That(Read(wrongMime), Is.EqualTo(bytes));
            Assert.That(presentation.GetPictureData().Count, Is.EqualTo(3));
            using var reopened = Open(Save(presentation));
            Assert.That(Read(reopened.Package.GetPart(occupied.PartName)), Is.EqualTo(bytes));
            Assert.That(reopened.GetPictureData().Count, Is.EqualTo(3));
        }

        [Test]
        public void UnreferencedImagePartsAreReusedAndReturnedDataDoesNotExposeInternalBuffers()
        {
            using var presentation = new XMLSlideShow();
            var bytes = Image(PictureType.PNG);
            var part = Part(presentation, "/custom/orphan.png", "image/png", bytes);
            var picture = presentation.AddPicture(bytes, PictureType.PNG);
            Assert.That(picture.GetPackagePart(), Is.SameAs(part));
            bytes[0] = 0;
            var copy = picture.Data;
            copy[0] = 0;
            Assert.That(picture.Data[0], Is.EqualTo(137));
            presentation.CreateSlide().CreatePicture(picture);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetPictureData().Single().GetPackagePart().PartName.Name,
                Is.EqualTo("/custom/orphan.png"));
        }

        [Test]
        public void PictureRelationshipsAreReusedByTypeRatherThanAnUnrelatedTargetMatch()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var picture = presentation.AddPicture(Image(PictureType.PNG), PictureType.PNG);
            var unrelated = slide.GetPackagePart().AddRelationship(picture.GetPackagePart().PartName,
                TargetMode.Internal, R + "/hyperlink");
            var first = slide.CreatePicture(picture);
            var second = slide.CreatePicture(picture);
            var id = Blip(first).GetAttribute("embed", R);
            Assert.That(id, Is.Not.EqualTo(unrelated.Id));
            Assert.That(slide.GetRelationById(id), Is.SameAs(picture));
            Assert.That(slide.GetRelations(), Does.Contain(picture));
            Assert.That(Blip(second).GetAttribute("embed", R), Is.EqualTo(id));
            Assert.That(slide.GetPackagePart().GetRelationshipsByType(PackageRelationshipTypes.IMAGE_PART).Size,
                Is.EqualTo(1));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetPackagePart().GetRelationship(unrelated.Id).RelationshipType,
                Is.EqualTo(R + "/hyperlink"));
            Assert.That(reopened.GetSlides()[0].GetShapes().Cast<XSLFPictureShape>()
                .Select(shape => shape.GetPictureData().Data), Is.All.EqualTo(picture.Data));
        }

        [TestCase("../media/image1.png")]
        [TestCase("../MEDIA/IMAGE1.PNG")]
        public void RelativeExistingImageRelationshipsAreReused(string target)
        {
            using var original = new XMLSlideShow();
            var slide = original.CreateSlide();
            var picture = original.AddPicture(Image(PictureType.PNG), PictureType.PNG);
            var relationship = slide.GetPackagePart().AddRelationship(
                new Uri(target, UriKind.Relative), TargetMode.Internal, PackageRelationshipTypes.IMAGE_PART);
            using var reopened = Open(Save(original));
            var created = reopened.GetSlides()[0].CreatePicture(reopened.GetPictureData().Single());
            Assert.That(Blip(created).GetAttribute("embed", R), Is.EqualTo(relationship.Id));
            Assert.That(reopened.GetSlides()[0].GetPackagePart()
                .GetRelationshipsByType(PackageRelationshipTypes.IMAGE_PART).Size, Is.EqualTo(1));
            Assert.That(reopened.GetSlides()[0].GetRelationById(relationship.Id), Is.InstanceOf<XSLFPictureData>());
        }

        [TestCase(PictureType.PNG)]
        [TestCase(PictureType.JPEG)]
        public void StreamsReadFromCurrentPositionAndRemainOwnedByTheCaller(PictureType type)
        {
            using var presentation = new XMLSlideShow();
            var bytes = Image(type);
            using var stream = new MemoryStream();
            stream.WriteByte(42);
            stream.Write(bytes, 0, bytes.Length);
            stream.Position = 1;
            var picture = presentation.AddPicture(stream, type);
            Assert.That(picture.Data, Is.EqualTo(bytes));
            Assert.That(stream.CanRead, Is.True);
            Assert.That(stream.Position, Is.EqualTo(stream.Length));
            using var nonSeekable = new NonSeekableStream(bytes);
            Assert.That(presentation.AddPicture(nonSeekable, type), Is.SameAs(picture));
            Assert.That(nonSeekable.CanRead, Is.True);
        }

        [TestCase(PictureType.PNG)]
        [TestCase(PictureType.JPEG)]
        public void InvalidEmptyTruncatedAndMismatchedImagesDoNotChangeThePackage(PictureType type)
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            var bytes = Image(type);
            var before = State(presentation);
            Assert.Throws<ArgumentNullException>(() => presentation.AddPicture((byte[])null, type));
            Assert.Throws<ArgumentException>(() => presentation.AddPicture(Array.Empty<byte>(), type));
            Assert.Throws<ArgumentException>(() => presentation.AddPicture(new byte[] { 1, 2, 3 }, type));
            Assert.Throws<ArgumentException>(() => presentation.AddPicture(bytes.Take(bytes.Length / 2).ToArray(), type));
            var other = type == PictureType.PNG ? PictureType.JPEG : PictureType.PNG;
            Assert.Throws<ArgumentException>(() => presentation.AddPicture(bytes, other));
            using var invalid = new MemoryStream(new byte[] { 1, 2 });
            Assert.Throws<ArgumentException>(() => presentation.AddPicture(invalid, type));
            Assert.That(invalid.CanRead, Is.True);
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [TestCase(PictureType.PNG)]
        [TestCase(PictureType.JPEG)]
        public void TruncatedImageTerminatorsDoNotCountAsCompleteImages(PictureType type)
        {
            using var presentation = new XMLSlideShow();
            var bytes = Image(type);
            var before = State(presentation);
            Assert.Throws<ArgumentException>(() => presentation.AddPicture(bytes.Take(bytes.Length - 1).ToArray(), type));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void StreamReadFailuresAreSurfacedWithoutAddingPartsOrClosingTheInput()
        {
            using var presentation = new XMLSlideShow();
            using var input = new FailingStream();
            var before = State(presentation);
            Assert.Throws<IOException>(() => presentation.AddPicture(input, PictureType.PNG));
            Assert.That(input.CanRead, Is.True);
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void InvalidExistingImageDataCannotCreateNewPictureShapes()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            Part(presentation, "/ppt/media/bad.png", "image/png", new byte[] { 1, 2, 3 });
            var data = presentation.GetPictureData().Single();
            var before = State(presentation);
            Assert.Throws<ArgumentException>(() => data.GetImageDimensionInPixels());
            Assert.Throws<ArgumentException>(() => slide.CreatePicture(data));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [TestCase(PictureType.GIF)]
        [TestCase(PictureType.EMF)]
        [TestCase(PictureType.None)]
        [TestCase((PictureType)int.MaxValue)]
        public void UnsupportedFormatsFailBeforeConsumingTheStream(PictureType type)
        {
            using var presentation = new XMLSlideShow();
            using var input = new MemoryStream(Image(PictureType.PNG));
            var before = State(presentation);
            Assert.Throws<NotSupportedException>(() => presentation.AddPicture(input, type));
            Assert.That(input.Position, Is.Zero);
            Assert.Throws<NotSupportedException>(() => presentation.AddPicture(input.ToArray(), type));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void NullUnreadableForeignAndRemovedInputsFailBeforeMutation()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            using var foreign = new XMLSlideShow();
            var other = foreign.AddPicture(Image(PictureType.PNG), PictureType.PNG);
            var before = State(presentation);
            Assert.Throws<ArgumentNullException>(() => presentation.AddPicture((Stream)null, PictureType.PNG));
            using var closed = new MemoryStream();
            closed.Dispose();
            Assert.Throws<ArgumentException>(() => presentation.AddPicture(closed, PictureType.PNG));
            Assert.Throws<ArgumentNullException>(() => slide.CreatePicture(null));
            Assert.Throws<ArgumentException>(() => slide.CreatePicture(other));
            Assert.That(State(presentation), Is.EqualTo(before));
            var data = presentation.AddPicture(Image(PictureType.PNG), PictureType.PNG);
            presentation.Package.RemovePart(data.GetPackagePart().PartName);
            before = State(presentation);
            Assert.Throws<InvalidOperationException>(() => slide.CreatePicture(data));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExternalPicturesAndEmbeddedCachesAreReadableButNotEditable(bool cache)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var bytes = Image(PictureType.PNG);
            var shape = slide.CreatePicture(presentation.AddPicture(bytes, PictureType.PNG));
            shape.SetAnchor(SKRect.Create(1, 2, 3, 4));
            var link = slide.GetPackagePart().AddExternalRelationship(
                "https://example.invalid/image.png", PackageRelationshipTypes.IMAGE_PART);
            Blip(shape).SetAttribute("link", R, link.Id);
            if (!cache)
            {
                Blip(shape).RemoveAttribute("embed", R);
            }
            Assert.That(shape.IsExternalLinkedPicture, Is.True);
            Assert.That(shape.GetPictureLink().OriginalString, Is.EqualTo("https://example.invalid/image.png"));
            Assert.That(shape.GetPictureData()?.Data, cache ? Is.EqualTo(bytes) : Is.Null);
            var before = State(presentation);
            Assert.Throws<NotSupportedException>(() => shape.Name = "Changed");
            Assert.Throws<NotSupportedException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.That(State(presentation), Is.EqualTo(before));
            var xml = shape.GetXmlObject().OuterXml;
            using var reopened = Open(Save(presentation));
            var saved = (XSLFPictureShape)reopened.GetSlides()[0].GetShapes().Single();
            Assert.That(saved.GetXmlObject().OuterXml, Is.EqualTo(xml));
            Assert.That(saved.GetPictureLink().OriginalString, Is.EqualTo("https://example.invalid/image.png"));
            Assert.That(saved.GetPictureData()?.Data, cache ? Is.EqualTo(bytes) : Is.Null);
        }

        [Test]
        public void UnsupportedImagesPreserveBytesAndShapesWithoutPretendingToDecodeThem()
        {
            using var presentation = new XMLSlideShow();
            var shape = NewPicture(presentation);
            var bytes = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"/>");
            var part = Part(presentation, "/ppt/media/unsupported.svg", "image/svg+xml", bytes);
            var relationship = shape.GetSheet().GetPackagePart().AddRelationship(
                part.PartName, TargetMode.Internal, PackageRelationshipTypes.IMAGE_PART);
            Blip(shape).SetAttribute("embed", R, relationship.Id);
            var data = shape.GetPictureData();
            Assert.That(data.Data, Is.EqualTo(bytes));
            Assert.That(data.PictureType, Is.EqualTo(PictureType.None));
            Assert.That(data.MimeType, Is.EqualTo("image/svg+xml"));
            Assert.Throws<NotSupportedException>(() => data.GetImageDimensionInPixels());
            var before = State(presentation);
            Assert.Throws<NotSupportedException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.Throws<NotSupportedException>(() => shape.Name = "Changed");
            Assert.Throws<NotSupportedException>(() => shape.GetSheet().CreatePicture(data));
            Assert.That(State(presentation), Is.EqualTo(before));
            var xml = shape.GetXmlObject().OuterXml;
            using var reopened = Open(Save(presentation));
            var saved = (XSLFPictureShape)reopened.GetSlides()[0].GetShapes().Single();
            Assert.That(saved.GetXmlObject().OuterXml, Is.EqualTo(xml));
            Assert.That(saved.GetPictureData().Data, Is.EqualTo(bytes));
        }

        [TestCase("missing")]
        [TestCase("empty")]
        [TestCase("wrongType")]
        [TestCase("externalEmbed")]
        [TestCase("internalLink")]
        [TestCase("noReference")]
        public void InvalidPictureReferencesFailExplicitlyWithoutMutation(string kind)
        {
            using var presentation = new XMLSlideShow();
            var shape = NewPicture(presentation);
            var source = shape.GetSheet().GetPackagePart();
            var part = shape.GetPictureData().GetPackagePart();
            var blip = Blip(shape);
            switch (kind)
            {
                case "missing": blip.SetAttribute("embed", R, "missing"); break;
                case "empty": blip.SetAttribute("embed", R, ""); break;
                case "wrongType":
                    blip.SetAttribute("embed", R,
                        source.AddRelationship(part.PartName, TargetMode.Internal, R + "/hyperlink").Id);
                    break;
                case "externalEmbed":
                    blip.SetAttribute("embed", R, source.AddExternalRelationship(
                        "https://example.invalid/image.png", PackageRelationshipTypes.IMAGE_PART).Id);
                    break;
                case "internalLink":
                    blip.RemoveAttribute("embed", R);
                    blip.SetAttribute("link", R,
                        source.AddRelationship(part.PartName, TargetMode.Internal, PackageRelationshipTypes.IMAGE_PART).Id);
                    break;
                case "noReference": blip.RemoveAttribute("embed", R); break;
            }
            var before = State(presentation);
            Assert.Throws<POIXMLException>(() => shape.GetPictureData());
            Assert.Throws<POIXMLException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void MissingPartsAndNonImageTargetsAreNotReturnedAsPictureData()
        {
            using var presentation = new XMLSlideShow();
            var shape = NewPicture(presentation);
            var source = shape.GetSheet().GetPackagePart();
            var part = Part(presentation, "/ppt/media/not-image.bin", "application/octet-stream", new byte[] { 1 });
            var relationship = source.AddRelationship(part.PartName, TargetMode.Internal, PackageRelationshipTypes.IMAGE_PART);
            Blip(shape).SetAttribute("embed", R, relationship.Id);
            var before = State(presentation);
            Assert.Throws<NotSupportedException>(() => shape.GetPictureData());
            Assert.Throws<NotSupportedException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.That(State(presentation), Is.EqualTo(before));
            presentation.Package.RemovePart(part.PartName);
            before = State(presentation);
            Assert.Throws<ArgumentException>(() => shape.GetPictureData());
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void MalformedContainersAndDetachedWrappersRejectPicturesWithoutAddingRelationships(bool duplicateExtensions)
        {
            using var presentation = new XMLSlideShow();
            var shape = NewPicture(presentation);
            var slide = (XSLFSlide)shape.GetSheet();
            var data = shape.GetPictureData();
            var group = slide.CreateGroup();
            group.GetXmlObject().ParentNode.RemoveChild(group.GetXmlObject());
            var before = State(presentation);
            Assert.Throws<InvalidOperationException>(() => group.CreatePicture(data));
            Assert.That(State(presentation), Is.EqualTo(before));
            shape.GetXmlObject().ParentNode.RemoveChild(shape.GetXmlObject());
            before = State(presentation);
            Assert.Throws<InvalidOperationException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.That(State(presentation), Is.EqualTo(before));
            var tree = slide.GetXmlObject().cSld["spTree", P];
            if (duplicateExtensions)
            {
                tree.AppendChild(tree.OwnerDocument.CreateElement("p", "extLst", P));
                tree.AppendChild(tree.OwnerDocument.CreateElement("p", "extLst", P));
            }
            else
            {
                tree.RemoveChild(tree["grpSpPr", P]);
            }
            before = State(presentation);
            Assert.Throws<POIXMLException>(() => slide.CreatePicture(data));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void CropsEffectsExtensionsAndGroupTransformsSurviveSupportedAnchorEdits()
        {
            using var presentation = new XMLSlideShow();
            var group = presentation.CreateSlide().CreateGroup();
            group.SetAnchor(SKRect.Create(100, 200, 300, 400));
            group.SetInteriorAnchor(SKRect.Create(1, 2, 30, 40));
            var shape = group.CreatePicture(presentation.AddPicture(Image(PictureType.PNG), PictureType.PNG));
            var xml = shape.GetXmlObject();
            var fill = xml["blipFill", P];
            var crop = xml.OwnerDocument.CreateElement("a", "srcRect", A);
            crop.SetAttribute("l", "10000");
            fill.InsertAfter(crop, Blip(shape));
            var effect = xml.OwnerDocument.CreateElement("a", "alphaModFix", A);
            effect.SetAttribute("amt", "50000");
            Blip(shape).AppendChild(effect);
            var transform = xml["spPr", P]["xfrm", A];
            transform.SetAttribute("rot", "60000");
            transform.SetAttribute("flipH", "1");
            var originalFill = fill.OuterXml;
            var originalGroupProperties = group.GetXmlObject()["grpSpPr", P].OuterXml;
            var extensions = xml.OwnerDocument.CreateElement("p", "extLst", P);
            var extension = xml.OwnerDocument.CreateElement("p", "ext", P);
            extension.SetAttribute("uri", "urn:npoi:preserved");
            extension.AppendChild(xml.OwnerDocument.CreateElement("test", "opaque", "urn:npoi:test"));
            extensions.AppendChild(extension);
            xml.AppendChild(extensions);
            var originalExtensions = extensions.OuterXml;
            shape.SetAnchor(SKRect.Create(-3, 4, 5, 6));
            shape.Name = "Edited picture";
            using var reopened = Open(Save(presentation));
            var savedGroup = (XSLFGroupShape)reopened.GetSlides()[0].GetShapes().Single();
            var saved = (XSLFPictureShape)savedGroup.GetShapes().Single();
            Assert.That(saved.GetXmlObject()["blipFill", P].OuterXml, Is.EqualTo(originalFill));
            Assert.That(savedGroup.GetXmlObject()["grpSpPr", P].OuterXml, Is.EqualTo(originalGroupProperties));
            Assert.That(saved.GetXmlObject()["spPr", P]["xfrm", A].GetAttribute("rot"), Is.EqualTo("60000"));
            Assert.That(saved.GetXmlObject()["spPr", P]["xfrm", A].GetAttribute("flipH"), Is.EqualTo("1"));
            Assert.That(saved.GetXmlObject()["extLst", P].OuterXml, Is.EqualTo(originalExtensions));
            Assert.That(saved.GetAnchor(), Is.EqualTo(SKRect.Create(-3, 4, 5, 6)));
        }

        [Test]
        public void ReadOnlyAndDisposedPresentationsRejectPictureChanges()
        {
            using var original = new XMLSlideShow();
            var shape = NewPicture(original);
            using var package = OPCPackage.Open(new MemoryStream(Save(original)), true);
            using var readOnly = new XMLSlideShow(package);
            var picture = readOnly.GetPictureData().Single();
            var saved = (XSLFPictureShape)readOnly.GetSlides()[0].GetShapes().Single();
            Assert.That(picture.GetImageDimensionInPixels(), Is.EqualTo(new SKSizeI(3, 2)));
            var before = State(readOnly);
            using var stream = new MemoryStream(picture.Data);
            Assert.Throws<InvalidOperationException>(() => readOnly.AddPicture(stream, PictureType.PNG));
            Assert.That(stream.Position, Is.Zero);
            Assert.Throws<InvalidOperationException>(() => readOnly.GetSlides()[0].CreatePicture(picture));
            Assert.Throws<InvalidOperationException>(() => saved.SetAnchor(SKRect.Empty));
            Assert.That(State(readOnly), Is.EqualTo(before));
            var originalData = shape.GetPictureData();
            original.Dispose();
            Assert.Throws<ObjectDisposedException>(() => original.GetPictureData());
            Assert.Throws<ObjectDisposedException>(() => original.AddPicture(stream, PictureType.PNG));
            Assert.Throws<ObjectDisposedException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.Throws<ObjectDisposedException>(() => _ = originalData.Data);
        }

        [Test]
        public void PicturesAreInsertedBeforeContainerExtensionsAndMediaSurvivesCommitTraversal()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var tree = slide.GetXmlObject().cSld["spTree", P];
            tree.AppendChild(tree.OwnerDocument.CreateElement("p", "extLst", P));
            var bytes = Image(PictureType.PNG);
            var data = presentation.AddPicture(bytes, PictureType.PNG);
            data.Committed = false;
            slide.CreatePicture(data);
            Assert.That(tree.LastChild.LocalName, Is.EqualTo("extLst"));
            Assert.That(tree.LastChild.PreviousSibling.LocalName, Is.EqualTo("pic"));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetPictureData().Single().Data, Is.EqualTo(bytes));
        }

        [TestCase("sample.pptx")]
        [TestCase("shapes.pptx")]
        public void RealFilesRetainExistingDependencyBytesWhenAddingPictures(string file)
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(file);
            using var presentation = new XMLSlideShow(input);
            var slide = presentation.GetSlides()[0];
            var originals = slide.GetShapes().Select(shape => shape.GetXmlObject().OuterXml).ToArray();
            var dependencies = presentation.Package.GetParts().Where(part => !part.IsRelationshipPart &&
                part != slide.GetPackagePart() && part.ContentType != XSLFRelation.MAIN.ContentType &&
                !part.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(part => part.PartName.Name, Read);
            using var image = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("clock.jpg");
            var picture = presentation.AddPicture(image, PictureType.JPEG);
            slide.CreatePicture(picture).SetAnchor(SKRect.Create(20, 30, 100, 100));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetShapes().Take(originals.Length)
                .Select(shape => shape.GetXmlObject().OuterXml), Is.EqualTo(originals));
            foreach (var dependency in dependencies)
            {
                Assert.That(Read(reopened.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key))),
                    Is.EqualTo(dependency.Value), dependency.Key);
            }
            Assert.That(((XSLFPictureShape)reopened.GetSlides()[0].GetShapes().Last()).GetPictureData().Data,
                Is.EqualTo(picture.Data));
        }

        [TestCase("shapes.pptx")]
        public void RealExistingEmbeddedPicturesCanMoveWithoutChangingOtherContent(string file)
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(file);
            using var presentation = new XMLSlideShow(input);
            var shape = presentation.GetSlides().SelectMany(AllShapes).OfType<XSLFPictureShape>()
                .First(item => !item.IsExternalLinkedPicture &&
                    (item.GetPictureData().PictureType == PictureType.PNG ||
                     item.GetPictureData().PictureType == PictureType.JPEG));
            var sheet = shape.GetSheet();
            var partName = sheet.GetPackagePart().PartName;
            var originalBytes = shape.GetPictureData().Data;
            var originalFill = shape.GetXmlObject()["blipFill", P].OuterXml;
            var others = AllShapes(sheet).Where(item => !ReferenceEquals(item.GetXmlObject(), shape.GetXmlObject()) &&
                !(item is XSLFGroupShape)).Select(item => item.GetXmlObject().OuterXml).ToArray();
            shape.SetAnchor(SKRect.Create(25, 35, 150, 100));
            using var reopened = Open(Save(presentation));
            var savedSheet = reopened.GetSlides().Single(item => item.GetPackagePart().PartName.Equals(partName));
            var saved = AllShapes(savedSheet).OfType<XSLFPictureShape>().Single(item => item.ID == shape.ID);
            Assert.That(saved.GetAnchor(), Is.EqualTo(SKRect.Create(25, 35, 150, 100)));
            Assert.That(saved.GetXmlObject()["blipFill", P].OuterXml, Is.EqualTo(originalFill));
            Assert.That(saved.GetPictureData().Data, Is.EqualTo(originalBytes));
            Assert.That(AllShapes(savedSheet).Where(item => item.ID != shape.ID && !(item is XSLFGroupShape))
                .Select(item => item.GetXmlObject().OuterXml), Is.EqualTo(others));
        }

        private static byte[] Image(PictureType type, SKColor? color = null)
        {
            using var bitmap = new SKBitmap(3, 2);
            bitmap.Erase(color ?? SKColors.CornflowerBlue);
            using var image = SKImage.FromBitmap(bitmap);
            using var encoded = image.Encode(
                type == PictureType.PNG ? SKEncodedImageFormat.Png : SKEncodedImageFormat.Jpeg, 90);
            return encoded.ToArray();
        }

        private static XSLFPictureShape NewPicture(XMLSlideShow presentation)
        {
            return presentation.CreateSlide().CreatePicture(presentation.AddPicture(Image(PictureType.PNG), PictureType.PNG));
        }

        private static IXSLFShapeContainer Container(XMLSlideShow presentation, string kind)
        {
            if (kind == "master")
                return presentation.GetSlideMasters()[0];
            if (kind == "layout")
                return presentation.GetSlideMasters()[0].GetSlideLayouts()[0];
            var slide = presentation.GetSlides().Count == 0 ? presentation.CreateSlide() : presentation.GetSlides()[0];
            if (kind == "slide")
                return slide;
            var group = slide.GetShapes().OfType<XSLFGroupShape>().FirstOrDefault() ?? slide.CreateGroup();
            return kind == "nested"
                ? group.GetShapes().OfType<XSLFGroupShape>().FirstOrDefault() ?? group.CreateGroup() : group;
        }

        private static IEnumerable<XSLFShape> AllShapes(IXSLFShapeContainer container)
        {
            foreach (var shape in container.GetShapes())
            {
                yield return shape;
                if (shape is XSLFGroupShape group)
                    foreach (var child in AllShapes(group))
                        yield return child;
            }
        }

        private static XmlElement Blip(XSLFPictureShape shape)
        {
            return shape.GetXmlObject()["blipFill", P]["blip", A];
        }

        private static PackagePart Part(XMLSlideShow presentation, string name, string mime, byte[] data)
        {
            var part = presentation.Package.CreatePart(PackagingUriHelper.CreatePartName(name), mime);
            using var output = part.GetOutputStream();
            output.Write(data, 0, data.Length);
            return part;
        }

        private static string State(XMLSlideShow presentation)
        {
            var parts = presentation.Package.GetParts().Where(part => !part.IsRelationshipPart)
                .Select(part => part.PartName.Name + "|" + part.ContentType + "|" +
                    string.Join(",", Relationships(part)));
            return string.Join("\n", parts) + string.Join("\n",
                presentation.GetSlides().Select(slide => slide.GetXmlObject().cSld.OuterXml));
        }

        private static IEnumerable<string> Relationships(PackagePart part)
        {
            foreach (var relationship in part.Relationships)
                yield return relationship.Id + ":" + relationship.RelationshipType + ":" +
                    relationship.TargetMode + ":" + relationship.TargetUri.OriginalString;
        }

        private static byte[] Read(PackagePart part)
        {
            using var input = part.GetInputStream();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] Save(XMLSlideShow presentation)
        {
            using var output = new MemoryStream();
            presentation.Write(output);
            return output.ToArray();
        }

        private static XMLSlideShow Open(byte[] bytes)
        {
            return new XMLSlideShow(new MemoryStream(bytes));
        }

        private sealed class NonSeekableStream : MemoryStream
        {
            internal NonSeekableStream(byte[] bytes) : base(bytes, false) { }

            public override bool CanSeek => false;
            public override long Length => throw new NotSupportedException();
            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                throw new NotSupportedException();
            }
        }

        private sealed class FailingStream : MemoryStream
        {
            public override int Read(byte[] buffer, int offset, int count)
            {
                throw new IOException("Image read failure.");
            }
        }
    }
}

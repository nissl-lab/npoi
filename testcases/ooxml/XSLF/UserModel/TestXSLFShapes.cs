using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.Util;
using NPOI.XSLF.UserModel;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFShapes
    {
        private const string P = CT_Presentation.Namespace;
        private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        [TestCase("slide")]
        [TestCase("layout")]
        [TestCase("master")]
        public void SheetsExposeOrderedReadOnlyShapeSnapshotsAndSaveEdits(string kind)
        {
            using var presentation = new XMLSlideShow();
            var sheet = GetSheet(presentation, kind);
            var empty = sheet.GetShapes();
            var rectangle = sheet.CreateAutoShape();
            var ellipse = sheet.CreateAutoShape(ST_ShapeType.ellipse);
            rectangle.Name = "Rectangle \u4e2d\u6587";
            rectangle.SetAnchor(SKRect.Create(12, 24, 180, 90));
            Assert.That(empty, Is.Empty);
            Assert.That(((ICollection<XSLFShape>)sheet.GetShapes()).IsReadOnly, Is.True);
            Assert.That(sheet.GetShapes().Select(shape => shape.ID), Is.EqualTo(new[] { rectangle.ID, ellipse.ID }));
            Assert.That(rectangle.Parent, Is.SameAs(sheet));
            Assert.That(rectangle.GetSheet(), Is.SameAs(sheet));
            Assert.That(rectangle.ShapeType, Is.EqualTo(ST_ShapeType.rect));
            using var reopened = Open(Save(presentation));
            var shapes = GetSheet(reopened, kind).GetShapes();
            Assert.That(shapes[0].Name, Is.EqualTo(rectangle.Name));
            Assert.That(shapes[0].ID, Is.EqualTo(rectangle.ID));
            Assert.That(shapes[0].GetAnchor(), Is.EqualTo(rectangle.GetAnchor()));
            Assert.That(((XSLFSimpleShape)shapes[1]).ShapeType, Is.EqualTo(ST_ShapeType.ellipse));
            shapes[0].Name = "Second save";
            using var again = Open(Save(reopened));
            Assert.That(GetSheet(again, kind).GetShapes()[0].Name, Is.EqualTo("Second save"));
            Assert.That(GetSheet(again, kind).GetShapes()[0].ID, Is.EqualTo(rectangle.ID));
        }

        [Test]
        public void NestedGroupsShareTheSheetIdSpaceAndKeepIndependentTransforms()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var simple = slide.CreateAutoShape();
            var group = slide.CreateGroup();
            group.SetAnchor(SKRect.Create(100, 200, 400, 300));
            group.SetInteriorAnchor(SKRect.Create(-10, -20, 200, 150));
            var nested = group.CreateGroup();
            nested.SetAnchor(SKRect.Create(10, 20, 80, 60));
            nested.SetInteriorAnchor(SKRect.Create(0, 0, 40, 30));
            var child = nested.CreateAutoShape(ST_ShapeType.triangle);
            child.SetAnchor(SKRect.Create(2, 3, 20, 10));
            var sibling = slide.CreateAutoShape();
            Assert.That(child.Parent, Is.SameAs(nested));
            Assert.That(child.GetSheet(), Is.SameAs(slide));
            var ids = new[] { simple.ID, group.ID, nested.ID, child.ID, sibling.ID };
            Assert.That(ids.Distinct().Count(), Is.EqualTo(ids.Length));
            Assert.That(ids, Does.Not.Contain(1U), "The shape-tree root identifier is reserved.");
            var transform = group.GetXmlObject()["grpSpPr", P]["xfrm", A];
            transform.SetAttribute("rot", "5400000");
            transform.SetAttribute("flipH", "1");
            transform.SetAttribute("flipV", "true");
            transform.SetAttribute("keep", "urn:npoi:test", "transform");
            var childXml = nested.GetXmlObject().OuterXml;
            group.SetAnchor(SKRect.Create(90, 80, 500, 600));
            Assert.That(group.GetInteriorAnchor(), Is.EqualTo(SKRect.Create(-10, -20, 200, 150)));
            group.SetInteriorAnchor(SKRect.Create(-30, -40, 250, 300));
            Assert.That(group.GetAnchor(), Is.EqualTo(SKRect.Create(90, 80, 500, 600)));
            Assert.That(nested.GetXmlObject().OuterXml, Is.EqualTo(childXml));
            Assert.That(transform.GetAttribute("rot"), Is.EqualTo("5400000"));
            Assert.That(transform.GetAttribute("flipH"), Is.EqualTo("1"));
            Assert.That(transform.GetAttribute("flipV"), Is.EqualTo("true"));
            Assert.That(transform.GetAttribute("keep", "urn:npoi:test"), Is.EqualTo("transform"));
            using var reopened = Open(Save(presentation));
            var saved = (XSLFGroupShape)reopened.GetSlides()[0].GetShapes()[1];
            Assert.That(saved.GetAnchor(), Is.EqualTo(group.GetAnchor()));
            Assert.That(saved.GetInteriorAnchor(), Is.EqualTo(group.GetInteriorAnchor()));
            Assert.That(AllShapes(reopened.GetSlides()[0]).Select(shape => shape.ID),
                Is.EqualTo(new[] { simple.ID, group.ID, nested.ID, child.ID, sibling.ID }));
        }

        [TestCase(0f, 0f, 0f, 0f)]
        [TestCase(-12.5f, 20.25f, 180.5f, 90.125f)]
        [TestCase(200000f, -200000f, 50000f, 100000f)]
        public void AnchorsUseRounded64BitEmusAndPoints(float x, float y, float width, float height)
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            var anchor = SKRect.Create(x, y, width, height);
            shape.SetAnchor(anchor);
            var transform = shape.GetXmlObject()["spPr", P]["xfrm", A];
            Assert.That(transform["off", A].GetAttribute("x"), Is.EqualTo(Emu(anchor.Left)));
            Assert.That(transform["off", A].GetAttribute("y"), Is.EqualTo(Emu(anchor.Top)));
            Assert.That(transform["ext", A].GetAttribute("cx"), Is.EqualTo(Emu(anchor.Width)));
            Assert.That(transform["ext", A].GetAttribute("cy"), Is.EqualTo(Emu(anchor.Height)));
            using var reopened = Open(Save(presentation));
            var saved = reopened.GetSlides()[0].GetShapes()[0].GetAnchor();
            Assert.That(saved.Left, Is.EqualTo(x).Within(0.001));
            Assert.That(saved.Top, Is.EqualTo(y).Within(0.001));
            Assert.That(saved.Width, Is.EqualTo(width).Within(0.001));
            Assert.That(saved.Height, Is.EqualTo(height).Within(0.001));
        }

        [TestCase(float.NaN)]
        [TestCase(float.PositiveInfinity)]
        [TestCase(float.NegativeInfinity)]
        [TestCase(float.MaxValue)]
        [TestCase(3000000000f)]
        public void NonfiniteAndOutOfRangeCoordinatesDoNotChangeXml(float invalid)
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            var before = shape.GetXmlObject().OuterXml;
            var anchors = new[]
            {
                new SKRect(invalid, 0, invalid, 10),
                new SKRect(0, invalid, 10, invalid),
                new SKRect(0, 0, invalid, 10),
                new SKRect(0, 0, 10, invalid)
            };
            foreach (var anchor in anchors)
            {
                Assert.Throws<ArgumentOutOfRangeException>(() => shape.SetAnchor(anchor));
                Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            }
        }

        [TestCase(false)]
        [TestCase(true)]
        public void NegativeExtentsFailAtomicallyForExteriorAndInteriorAnchors(bool interior)
        {
            using var presentation = new XMLSlideShow();
            var group = presentation.CreateSlide().CreateGroup();
            var before = group.GetXmlObject().OuterXml;
            foreach (var rectangle in new[] { new SKRect(0, 0, -1, 10), new SKRect(0, 0, 10, -1) })
            {
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    if (interior)
                        group.SetInteriorAnchor(rectangle);
                    else
                        group.SetAnchor(rectangle);
                });
                Assert.That(group.GetXmlObject().OuterXml, Is.EqualTo(before));
            }
        }

        [TestCase("-1", "ext", "cx")]
        [TestCase("2147483648", "ext", "cx")]
        [TestCase("27273042316901", "ext", "cy")]
        [TestCase("-27273042329601", "off", "x")]
        [TestCase("1in", "off", "y")]
        [TestCase("bad", "off", "x")]
        public void InvalidOrUnsupportedStoredCoordinatesFailExplicitly(string value, string element, string attribute)
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            shape.GetXmlObject()["spPr", P]["xfrm", A][element, A].SetAttribute(attribute, value);
            Assert.Throws<POIXMLException>(() => shape.GetAnchor());
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExtentsUseTheirOwnLimitRatherThanTheWiderPositionRange(bool interior)
        {
            using var presentation = new XMLSlideShow();
            var group = presentation.CreateSlide().CreateGroup();
            var before = group.GetXmlObject().OuterXml;
            foreach (var rectangle in new[] { SKRect.Create(0, 0, 169094, 10), SKRect.Create(0, 0, 10, 169094) })
            {
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    if (interior)
                        group.SetInteriorAnchor(rectangle);
                    else
                        group.SetAnchor(rectangle);
                });
                Assert.That(group.GetXmlObject().OuterXml, Is.EqualTo(before));
            }
            var valid = SKRect.Create(200000, -200000, 169093, 100);
            if (interior)
            {
                group.SetInteriorAnchor(valid);
                Assert.That(group.GetInteriorAnchor(), Is.EqualTo(valid));
            }
            else
            {
                group.SetAnchor(valid);
                Assert.That(group.GetAnchor(), Is.EqualTo(valid));
            }
        }

        [Test]
        public void ExactStoredBoundsAndAdjacentFloatExtentsUseTheCorrectLimits()
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            var transform = shape.GetXmlObject()["spPr", P]["xfrm", A];
            transform["off", A].SetAttribute("x", "-27273042329600");
            transform["off", A].SetAttribute("y", "27273042316900");
            var before = shape.GetXmlObject().OuterXml;
            Assert.That(shape.GetAnchor().Left, Is.EqualTo((float)Units.ToPoints(-27273042329600L)));
            Assert.That(shape.GetAnchor().Top, Is.EqualTo((float)Units.ToPoints(27273042316900L)));
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            transform["off", A].SetAttribute("x", "0");
            transform["off", A].SetAttribute("y", "0");
            transform["ext", A].SetAttribute("cx", "2147483647");
            Assert.That(shape.GetAnchor().Width, Is.EqualTo((float)Units.ToPoints(int.MaxValue)));
            shape.SetAnchor(SKRect.Create(0, 0, 169093.1875f, 0));
            Assert.That(transform["ext", A].GetAttribute("cx"), Is.EqualTo("2147483481"));
            before = shape.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => shape.SetAnchor(SKRect.Create(0, 0, 169093.203125f, 0)));
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void MissingTransformsAreNotInventedByGettersAndCanBeSetExplicitlyInSchemaOrder()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateAutoShape();
            var properties = shape.GetXmlObject()["spPr", P];
            properties.RemoveChild(properties["xfrm", A]);
            var before = shape.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => shape.GetAnchor());
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            shape.SetAnchor(SKRect.Create(1, 2, 3, 4));
            Assert.That(properties.FirstChild.LocalName, Is.EqualTo("xfrm"));
            Assert.That(shape.GetAnchor(), Is.EqualTo(SKRect.Create(1, 2, 3, 4)));
            var group = slide.CreateGroup();
            properties = group.GetXmlObject()["grpSpPr", P];
            properties.RemoveAll();
            Assert.Throws<NotSupportedException>(() => group.GetInteriorAnchor());
            group.SetInteriorAnchor(SKRect.Create(0, 0, 10, 20));
            Assert.Throws<NotSupportedException>(() => group.GetAnchor());
            group.SetAnchor(SKRect.Create(20, 30, 40, 50));
            Assert.That(properties["xfrm", A].ChildNodes.Cast<XmlNode>().Select(node => node.LocalName),
                Is.EqualTo(new[] { "off", "ext", "chOff", "chExt" }));
        }

        [Test]
        public void OpaqueAndAlternateContentStayInOrderAndReserveAllTheirShapeIds()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var tree = Tree(slide);
            Append(tree, "<p:pic><p:nvPicPr><p:cNvPr id=\"2\" name=\"Opaque picture\"/></p:nvPicPr>" +
                "<p:spPr/><a:unfamiliar value=\"retain\"/></p:pic>");
            Append(tree, "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" +
                "<mc:Choice Requires=\"a\"><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Choice\"/></p:nvSpPr></p:sp></mc:Choice>" +
                "<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Fallback\"/></p:nvSpPr></p:sp></mc:Fallback>" +
                "</mc:AlternateContent>");
            Append(tree, "<p:extLst><p:ext uri=\"urn:npoi:test\"><a:unfamiliar/></p:ext></p:extLst>");
            var originals = slide.GetShapes().Select(item => item.GetXmlObject().OuterXml).ToArray();
            var shape = slide.CreateAutoShape();
            Assert.That(shape.ID, Is.EqualTo(4));
            Assert.That(tree.LastChild.LocalName, Is.EqualTo("extLst"));
            Assert.That(slide.GetShapes().Select(item => item.GetXmlObject().LocalName),
                Is.EqualTo(new[] { "pic", "AlternateContent", "sp" }));
            var picture = slide.GetShapes()[0];
            Assert.That(picture, Is.TypeOf<XSLFUnsupportedShape>());
            Assert.That(picture.ID, Is.EqualTo(2));
            Assert.That(picture.Name, Is.EqualTo("Opaque picture"));
            Assert.Throws<NotSupportedException>(() => picture.Name = "Not supported");
            Assert.Throws<NotSupportedException>(() => picture.GetAnchor());
            Assert.Throws<NotSupportedException>(() => picture.SetAnchor(SKRect.Empty));
            Assert.Throws<NotSupportedException>(() => _ = slide.GetShapes()[1].ID);
            shape.SetAnchor(SKRect.Create(12, 34, 56, 78));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetShapes().Take(2).Select(item => item.GetXmlObject().OuterXml),
                Is.EqualTo(originals));
        }

        [Test]
        public void EditingASimpleShapePreservesItsTextStyleAndUnknownTransformContent()
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            Append(shape.GetXmlObject(), "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b=\"1\"/>" +
                "<a:t>Existing text</a:t></a:r></a:p></p:txBody>");
            var properties = shape.GetXmlObject()["spPr", P];
            Append(properties, "<a:solidFill><a:srgbClr val=\"ABCDEF\"/></a:solidFill>");
            var transform = properties["xfrm", A];
            transform.SetAttribute("rot", "60000");
            Append(transform, "<a:extLst><a:ext uri=\"urn:npoi:test\"><p:unknown/></a:ext></a:extLst>");
            var text = shape.GetXmlObject()["txBody", P].OuterXml;
            var fill = properties["solidFill", A].OuterXml;
            var extension = transform["extLst", A].OuterXml;
            shape.Name = "Renamed";
            shape.SetAnchor(SKRect.Create(1, 2, 3, 4));
            using var reopened = Open(Save(presentation));
            var saved = reopened.GetSlides()[0].GetShapes()[0].GetXmlObject();
            Assert.That(saved["txBody", P].OuterXml, Is.EqualTo(text));
            Assert.That(saved["spPr", P]["solidFill", A].OuterXml, Is.EqualTo(fill));
            Assert.That(saved["spPr", P]["xfrm", A]["extLst", A].OuterXml, Is.EqualTo(extension));
            Assert.That(saved["spPr", P]["xfrm", A].GetAttribute("rot"), Is.EqualTo("60000"));
        }

        [TestCase("sample.pptx")]
        [TestCase("shapes.pptx")]
        public void RealFilesKeepOtherShapesAndDependencyBytesWhenEditingASupportedShape(string file)
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(file);
            using var presentation = new XMLSlideShow(input);
            var shape = presentation.GetSlides().SelectMany(AllShapes).OfType<XSLFSimpleShape>().First();
            var slide = shape.GetSheet();
            var originalIds = AllShapes(slide).Where(item => !(item is XSLFUnsupportedShape)).Select(item => item.ID).ToArray();
            var untouchedShapes = AllShapes(slide).Where(item => !ReferenceEquals(item.GetXmlObject(), shape.GetXmlObject()))
                .Where(item => !(item is XSLFGroupShape)).Select(item => item.GetXmlObject().OuterXml).ToArray();
            var dependencies = presentation.Package.GetParts().Where(part => !part.IsRelationshipPart &&
                part != slide.GetPackagePart() && part.ContentType != XSLFRelation.MAIN.ContentType &&
                !part.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(part => part.PartName.Name, ReadBytes);
            var relationships = PartRelationships(presentation.Package).ToArray();
            var name = slide.GetPackagePart().PartName;
            shape.Name = "Edited existing shape";
            shape.SetAnchor(SKRect.Create(10, 20, 200, 100));
            using var reopened = Open(Save(presentation));
            var savedSlide = reopened.GetSlides().Single(item => item.GetPackagePart().PartName.Equals(name));
            var shapes = AllShapes(savedSlide).ToArray();
            Assert.That(shapes.Where(item => !(item is XSLFUnsupportedShape)).Select(item => item.ID), Is.EqualTo(originalIds));
            Assert.That(shapes.Single(item => item.Name == "Edited existing shape").GetAnchor(),
                Is.EqualTo(SKRect.Create(10, 20, 200, 100)));
            Assert.That(shapes.Where(item => item.ID != shape.ID && !(item is XSLFGroupShape))
                .Select(item => item.GetXmlObject().OuterXml), Is.EqualTo(untouchedShapes));
            foreach (var dependency in dependencies)
            {
                Assert.That(ReadBytes(reopened.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key))),
                    Is.EqualTo(dependency.Value), dependency.Key);
            }
            Assert.That(PartRelationships(reopened.Package), Is.EqualTo(relationships));
        }

        [TestCase(ST_ShapeType.none)]
        [TestCase(ST_ShapeType.lineInv)]
        [TestCase((ST_ShapeType)int.MaxValue)]
        public void InvalidPresetsFailWithoutChangingTheContainer(ST_ShapeType preset)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var before = Tree(slide).OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => slide.CreateAutoShape(preset));
            Assert.That(Tree(slide).OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void FullUnsignedIdentifiersArePreservedAndInvalidIdentifiersRejectCreation()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateAutoShape();
            var properties = shape.GetXmlObject()["nvSpPr", P]["cNvPr", P];
            properties.SetAttribute("id", uint.MaxValue.ToString(CultureInfo.InvariantCulture));
            Assert.That(shape.ID, Is.EqualTo(uint.MaxValue));
            Assert.That(slide.CreateAutoShape().ID, Is.EqualTo(2));
            properties.SetAttribute("id", "4294967296");
            var before = Tree(slide).OuterXml;
            Assert.Throws<POIXMLException>(() => slide.CreateAutoShape());
            Assert.That(Tree(slide).OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void DuplicateTransformsAndMalformedContainersFailBeforeMutation()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateAutoShape();
            var properties = shape.GetXmlObject()["spPr", P];
            properties.AppendChild(properties["xfrm", A].CloneNode(true));
            var before = Tree(slide).OuterXml;
            Assert.Throws<POIXMLException>(() => shape.GetAnchor());
            Assert.Throws<POIXMLException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.That(Tree(slide).OuterXml, Is.EqualTo(before));
            Tree(slide).AppendChild(Tree(slide)["grpSpPr", P].CloneNode(true));
            before = Tree(slide).OuterXml;
            Assert.Throws<POIXMLException>(() => slide.GetShapes());
            Assert.Throws<POIXMLException>(() => slide.CreateGroup());
            Assert.That(Tree(slide).OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void ReadonlyClosedAndDetachedShapesRejectMutations()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateAutoShape();
            var group = slide.CreateGroup();
            using var package = OPCPackage.Open(new MemoryStream(Save(presentation)), true);
            using var readOnly = new XMLSlideShow(package);
            var readOnlySlide = readOnly.GetSlides()[0];
            var before = Tree(readOnlySlide).OuterXml;
            Assert.Throws<InvalidOperationException>(() => readOnlySlide.CreateAutoShape());
            Assert.Throws<InvalidOperationException>(() => readOnlySlide.CreateGroup());
            Assert.Throws<InvalidOperationException>(() => readOnlySlide.GetShapes()[0].Name = "Readonly");
            Assert.Throws<InvalidOperationException>(() => readOnlySlide.GetShapes()[0].SetAnchor(SKRect.Empty));
            Assert.Throws<InvalidOperationException>(() => ((XSLFGroupShape)readOnlySlide.GetShapes()[1]).CreateGroup());
            Assert.That(Tree(readOnlySlide).OuterXml, Is.EqualTo(before));
            Tree(slide).RemoveChild(shape.GetXmlObject());
            Assert.Throws<InvalidOperationException>(() => shape.Name = "Detached");
            Assert.Throws<InvalidOperationException>(() => shape.SetAnchor(SKRect.Empty));
            presentation.RemoveSlide(0);
            Assert.Throws<InvalidOperationException>(() => group.CreateAutoShape());
            Assert.Throws<InvalidOperationException>(() => group.SetInteriorAnchor(SKRect.Empty));
            Assert.Throws<InvalidOperationException>(() => slide.CreateGroup());
            presentation.Dispose();
            Assert.Throws<ObjectDisposedException>(() => group.SetAnchor(SKRect.Empty));
        }

        [Test]
        public void DetachedGroupsAndReplacedTreesCannotBeEditedThroughOldWrappers()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var group = slide.CreateGroup();
            var child = group.CreateAutoShape();
            Tree(slide).RemoveChild(group.GetXmlObject());
            Assert.Throws<InvalidOperationException>(() => group.CreateGroup());
            Assert.Throws<InvalidOperationException>(() => child.Name = "Detached group child");
            var shape = slide.CreateAutoShape();
            var oldTree = Tree(slide);
            slide.GetXmlObject().cSld.ReplaceChild(oldTree.CloneNode(true), oldTree);
            Assert.Throws<InvalidOperationException>(() => shape.SetAnchor(SKRect.Empty));
            slide.CreateAutoShape();
            Assert.That(slide.GetShapes().Select(item => item.ID).Distinct().Count(), Is.EqualTo(2));
        }

        [Test]
        public void LowLevelPartRemovalAndCommittedSheetsRejectUnsavableChanges()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateAutoShape();
            slide.Committed = true;
            Assert.Throws<InvalidOperationException>(() => shape.Name = "Not saved");
            slide.Committed = false;
            presentation.Package.RemovePart(slide.GetPackagePart().PartName);
            Assert.Throws<InvalidOperationException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.Throws<InvalidOperationException>(() => slide.CreateGroup());
        }

        [Test]
        public void NamesAndGeometryGettersHaveNoSilentInvalidInputFallbacks()
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            var before = shape.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentNullException>(() => shape.Name = null);
            Assert.Throws<XmlException>(() => shape.Name = "\u0001");
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            var properties = shape.GetXmlObject()["spPr", P];
            properties["prstGeom", A].SetAttribute("prst", "invalid");
            Assert.Throws<POIXMLException>(() => _ = shape.ShapeType);
            properties.RemoveChild(properties["prstGeom", A]);
            Assert.That(shape.ShapeType, Is.Null, "A missing preset is not silently replaced with a rectangle.");
        }

        private static XSLFSheet GetSheet(XMLSlideShow presentation, string kind)
        {
            switch (kind)
            {
                case "master":
                    return presentation.GetSlideMasters()[0];
                case "layout":
                    return presentation.GetSlideMasters()[0].GetSlideLayouts()[0];
                default:
                    return presentation.GetSlides().FirstOrDefault() ?? presentation.CreateSlide();
            }
        }

        private static IEnumerable<XSLFShape> AllShapes(IXSLFShapeContainer container)
        {
            foreach (var shape in container.GetShapes())
            {
                yield return shape;
                if (shape is XSLFGroupShape group)
                {
                    foreach (var child in AllShapes(group))
                        yield return child;
                }
            }
        }

        private static XmlElement Tree(XSLFSlide slide)
        {
            return slide.GetXmlObject().cSld["spTree", P];
        }

        private static void Append(XmlElement parent, string xml)
        {
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml("<root xmlns:p=\"" + P + "\" xmlns:a=\"" + A + "\">" + xml + "</root>");
            parent.AppendChild(parent.OwnerDocument.ImportNode(document.DocumentElement.FirstChild, true));
        }

        private static string Emu(float points)
        {
            return Math.Round((double)points * Units.EMU_PER_POINT).ToString(CultureInfo.InvariantCulture);
        }

        private static XMLSlideShow Open(byte[] bytes)
        {
            return new XMLSlideShow(new MemoryStream(bytes));
        }

        private static byte[] ReadBytes(PackagePart part)
        {
            using var input = part.GetInputStream();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        private static IEnumerable<string> PartRelationships(OPCPackage package)
        {
            var relationships = new List<string>();
            foreach (var part in package.GetParts().Where(item => !item.IsRelationshipPart))
            {
                foreach (var relationship in part.Relationships)
                {
                    relationships.Add(part.PartName.Name + "|" + relationship.Id + "|" +
                        relationship.RelationshipType + "|" + relationship.TargetMode + "|" + relationship.TargetUri);
                }
            }
            return relationships.OrderBy(value => value, StringComparer.Ordinal);
        }

        private static byte[] Save(XMLSlideShow presentation)
        {
            using var output = new MemoryStream();
            presentation.Write(output);
            return output.ToArray();
        }
    }
}

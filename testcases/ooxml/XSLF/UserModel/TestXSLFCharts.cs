using System;
using System.Collections.Generic;
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
    public class TestXSLFCharts
    {
        private const string P = CT_Presentation.Namespace;
        private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string R = CT_Presentation.RelationshipsNamespace;

        [Test]
        public void ExistingChartsHaveDetachedXmlAndModelSnapshotsAndIndependentByteArrays()
        {
            using var presentation = Pie();
            var chart = presentation.GetCharts().Single();
            var original = chart.Data;
            var bytes = chart.Data;
            bytes[0] = 0;
            var xml = chart.GetXmlSnapshot();
            Assert.That(xml.NamespaceURI, Is.EqualTo(C));
            Assert.That(xml.LocalName, Is.EqualTo("chartSpace"));
            xml.SetAttribute("snapshotOnly", "yes");
            var model = chart.GetChartSpaceSnapshot();
            Assert.That(model.chart.plotArea.pieChart.Count, Is.EqualTo(1));
            Assert.That(model.chart.plotArea.pieChart[0].ser.Count, Is.EqualTo(1));
            model.chart.plotArea.pieChart[0].ser.Clear();
            Assert.That(chart.GetChartSpaceSnapshot().chart.plotArea.pieChart[0].ser.Count, Is.EqualTo(1));
            Assert.That(chart.GetXmlSnapshot().HasAttribute("snapshotOnly"), Is.False);
            Assert.That(chart.Data, Is.EqualTo(original));
            Assert.That(((ICollection<XSLFChart>)presentation.GetCharts()).IsReadOnly, Is.True);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCharts().Single().Data, Is.EqualTo(original));
        }

        [TestCase("slide")]
        [TestCase("layout")]
        [TestCase("master")]
        [TestCase("group")]
        [TestCase("nested")]
        public void AllContainersReuseChartPartsWithUniqueIdsAndEmptyAnchors(string kind)
        {
            using var presentation = Pie();
            var chart = presentation.GetCharts().Single();
            var container = Container(presentation, kind);
            var snapshot = container.GetShapes();
            var first = container.CreateChart(chart);
            var second = container.CreateChart(chart);
            Assert.That(first.Parent, Is.SameAs(container));
            Assert.That(first.ID, Is.Not.EqualTo(second.ID));
            Assert.That(first.GetAnchor(), Is.EqualTo(SKRect.Empty));
            Assert.That(first.GetChart(), Is.SameAs(chart));
            Assert.That(snapshot.Count, Is.EqualTo(container.GetShapes().Count - 2));
            Assert.That(first.GetSheet().GetPackagePart().GetRelationshipsByType(XSLFRelation.CHART.Relation).Size,
                Is.EqualTo(1));
            Assert.That(Reference(first).GetAttribute("id", R), Is.EqualTo(Reference(second).GetAttribute("id", R)));
            first.SetAnchor(SKRect.Create(-10, 20, 300, 200));
            first.Name = "Reused chart <&>";
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            var saved = Container(again, kind).GetShapes().OfType<XSLFChartShape>().Single(item => item.ID == first.ID);
            Assert.That(saved.Name, Is.EqualTo("Reused chart <&>"));
            Assert.That(saved.GetAnchor(), Is.EqualTo(SKRect.Create(-10, 20, 300, 200)));
            Assert.That(saved.GetChart(), Is.SameAs(again.GetCharts().Single()));
            Assert.That(saved.GetChart().Data, Is.EqualTo(chart.Data));
        }

        [Test]
        public void ChartAndWorkbookBytesAndRelationsSurviveFrameEditsAndSharedSlideRemoval()
        {
            using var presentation = Pie();
            var chart = presentation.GetCharts().Single();
            var dependencies = Dependencies(presentation);
            var first = presentation.GetSlides()[0];
            var native = AllShapes(first).OfType<XSLFChartShape>().Single();
            Assert.That(first.GetRelationById(Reference(native).GetAttribute("id", R)), Is.InstanceOf<XSLFChart>());
            native.Name = "Moved native chart";
            native.SetAnchor(SKRect.Create(40, 50, 360, 200));
            var second = presentation.CreateSlide();
            var reused = second.CreateGroup().CreateChart(chart);
            reused.SetAnchor(SKRect.Create(1, 2, 100, 80));
            presentation.RemoveSlide(0);
            Assert.Throws<InvalidOperationException>(() => native.SetAnchor(SKRect.Empty));
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            Assert.That(again.GetCharts().Count, Is.EqualTo(1));
            foreach (var dependency in dependencies)
            {
                var part = again.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key));
                Assert.That(Read(part), Is.EqualTo(dependency.Value.Bytes), dependency.Key);
                Assert.That(Relationships(part), Is.EqualTo(dependency.Value.Relationships), dependency.Key);
            }
            Assert.That(AllShapes(again.GetSlides().Last()).OfType<XSLFChartShape>().Single().GetChart().Data,
                Is.EqualTo(chart.Data));
        }

        [Test]
        public void ChartRelationshipsAreSelectedByTypeAndRemainRegisteredInTheDocumentGraph()
        {
            using var presentation = Pie();
            var chart = presentation.GetCharts().Single();
            var slide = presentation.CreateSlide();
            var unrelated = slide.GetPackagePart().AddRelationship(
                chart.GetPackagePart().PartName, TargetMode.Internal, R + "/hyperlink");
            var first = slide.CreateChart(chart);
            var second = slide.CreateChart(chart);
            var id = Reference(first).GetAttribute("id", R);
            Assert.That(id, Is.Not.EqualTo(unrelated.Id));
            Assert.That(Reference(second).GetAttribute("id", R), Is.EqualTo(id));
            Assert.That(slide.GetRelationById(id), Is.SameAs(chart));
            Assert.That(slide.GetRelations(), Does.Contain(chart));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Last().GetPackagePart().GetRelationship(unrelated.Id).RelationshipType,
                Is.EqualTo(R + "/hyperlink"));
            Assert.That(reopened.GetSlides().Last().GetShapes().Cast<XSLFChartShape>()
                .Select(shape => shape.GetChart()), Is.All.SameAs(reopened.GetCharts().Single()));
        }

        [TestCase("../charts/chart1.xml")]
        [TestCase("../CHARTS/CHART1.XML")]
        public void RelativeAndCaseEquivalentChartTargetsAreReused(string target)
        {
            using var presentation = Pie();
            var slide = presentation.CreateSlide();
            var relationship = slide.GetPackagePart().AddRelationship(
                new Uri(target, UriKind.Relative), TargetMode.Internal, XSLFRelation.CHART.Relation);
            using var reopened = Open(Save(presentation));
            var savedSlide = reopened.GetSlides().Last();
            var frame = savedSlide.CreateChart(reopened.GetCharts().Single());
            Assert.That(Reference(frame).GetAttribute("id", R), Is.EqualTo(relationship.Id));
            Assert.That(savedSlide.GetPackagePart().GetRelationshipsByType(XSLFRelation.CHART.Relation).Size, Is.EqualTo(1));
        }

        [Test]
        public void EmptyAndUnreferencedChartCollectionsAreReadOnlySnapshots()
        {
            using var presentation = new XMLSlideShow();
            var snapshot = presentation.GetCharts();
            Assert.That(snapshot, Is.Empty);
            using var source = Pie();
            var original = source.GetCharts().Single().Data;
            var part = Part(presentation, "/ppt/charts/orphan.xml", XSLFRelation.CHART.ContentType, original);
            var chart = presentation.GetCharts().Single();
            Assert.That(chart.GetPackagePart(), Is.SameAs(part));
            Assert.That(chart.GetXmlSnapshot().LocalName, Is.EqualTo("chartSpace"));
            Assert.That(snapshot, Is.Empty);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCharts().Single().Data, Is.EqualTo(original));
        }

        [Test]
        public void ChartContentTypesAreMatchedCaseInsensitively()
        {
            using var presentation = Pie();
            var part = Part(presentation, "/ppt/charts/uppercase.xml",
                XSLFRelation.CHART.ContentType.ToUpperInvariant(), presentation.GetCharts().Single().Data);
            var chart = presentation.GetCharts().Single(item => item.GetPackagePart() == part);
            var frame = presentation.CreateSlide().CreateChart(chart);
            using var reopened = Open(Save(presentation));
            var saved = (XSLFChartShape)reopened.GetSlides().Last().GetShapes().Single();
            Assert.That(saved.GetChart().Data, Is.EqualTo(chart.Data));
            Assert.That(reopened.GetSlides().Last().GetRelationById(Reference(saved).GetAttribute("id", R)),
                Is.InstanceOf<XSLFChart>());
        }

        [Test]
        public void ChartExPartsAndFramesRemainOpaqueAndBytePreserved()
        {
            using var presentation = Pie();
            const string chartExNamespace = "http://schemas.microsoft.com/office/drawing/2014/chartex";
            var bytes = Encoding.UTF8.GetBytes("<cx:chartSpace xmlns:cx=\"" + chartExNamespace +
                "\"><cx:unknown preserved=\"yes\"/></cx:chartSpace>");
            var part = Part(presentation, "/ppt/charts/chartEx1.xml", "application/vnd.ms-office.chartex+xml", bytes);
            var slide = presentation.GetSlides()[0];
            var frame = AllShapes(slide).OfType<XSLFChartShape>().Single();
            var data = frame.GetXmlObject()["graphic", A]["graphicData", A];
            data.RemoveAll();
            data.SetAttribute("uri", chartExNamespace);
            data.SetAttribute("xmlns:cx", chartExNamespace);
            data.SetAttribute("xmlns:r", R);
            var reference = data.OwnerDocument.CreateElement("cx", "chart", chartExNamespace);
            var id = data.OwnerDocument.CreateAttribute("r", "id", R);
            id.Value = slide.GetPackagePart().AddRelationship(part.PartName,
                TargetMode.Internal, "http://schemas.microsoft.com/office/2014/relationships/chartEx").Id;
            reference.Attributes.Append(id);
            data.AppendChild(reference);
            var xml = frame.GetXmlObject().OuterXml;
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            Assert.That(again.GetCharts().Count, Is.EqualTo(1));
            var saved = AllShapes(again.GetSlides()[0]).Single(item => item.ID == frame.ID);
            Assert.That(saved, Is.TypeOf<XSLFUnsupportedShape>());
            Assert.That(saved.GetXmlObject().OuterXml, Is.EqualTo(xml));
            Assert.That(Read(again.Package.GetPart(part.PartName)), Is.EqualTo(bytes));
            Assert.That(again.GetSlides()[0].GetRelationById(reference.GetAttribute("id", R)),
                Is.Not.InstanceOf<XSLFChart>());
        }

        [Test]
        public void UnknownChartExtensionsAndExternalWorkbookLinksArePreservedWithoutFetchingThem()
        {
            using var presentation = Pie();
            var chart = presentation.GetCharts().Single();
            var part = chart.GetPackagePart();
            var document = chart.GetXmlSnapshot().OwnerDocument;
            var root = document.DocumentElement;
            var extensions = root["extLst", C] ?? document.CreateElement("c", "extLst", C);
            if (extensions.ParentNode == null)
                root.AppendChild(extensions);
            var extension = document.CreateElement("c", "ext", C);
            extension.SetAttribute("uri", "urn:npoi:unknown-chart");
            var unknown = document.CreateElement("unknown", "payload", "urn:npoi:test");
            unknown.SetAttribute("preserved", "yes");
            extension.AppendChild(unknown);
            extensions.AppendChild(extension);
            var external = part.AddExternalRelationship("https://example.invalid/never-load.xlsx", R + "/package");
            var externalData = root["externalData", C];
            externalData.SetAttribute("id", R, external.Id);
            Write(part, Encoding.UTF8.GetBytes(document.OuterXml));
            Assert.Throws<InvalidOperationException>(() => chart.GetXmlSnapshot());
            chart = presentation.GetCharts().Single();
            part = chart.GetPackagePart();
            var before = Read(part);
            Assert.That(chart.GetXmlSnapshot().GetElementsByTagName("payload", "urn:npoi:test").Count, Is.EqualTo(1));
            Assert.That(chart.GetChartSpaceSnapshot().chart.plotArea.pieChart.Count, Is.EqualTo(1));
            presentation.CreateSlide().CreateChart(chart).SetAnchor(SKRect.Create(10, 20, 300, 200));
            chart.Committed = false;
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCharts().Single().Data, Is.EqualTo(before));
            var relationship = reopened.GetCharts().Single().GetPackagePart().GetRelationship(external.Id);
            Assert.That(relationship.TargetMode, Is.EqualTo(TargetMode.External));
            Assert.That(relationship.TargetUri.OriginalString, Is.EqualTo("https://example.invalid/never-load.xlsx"));
        }

        [TestCase("missing")]
        [TestCase("empty")]
        [TestCase("wrongType")]
        [TestCase("external")]
        [TestCase("wrongMime")]
        public void InvalidChartReferencesFailBeforeFrameMutation(string kind)
        {
            using var presentation = Pie();
            var shape = AllShapes(presentation.GetSlides()[0]).OfType<XSLFChartShape>().Single();
            var source = shape.GetSheet().GetPackagePart();
            var chart = shape.GetChart();
            switch (kind)
            {
                case "missing": Reference(shape).SetAttribute("id", R, "missing"); break;
                case "empty": Reference(shape).SetAttribute("id", R, ""); break;
                case "wrongType":
                    Reference(shape).SetAttribute("id", R,
                        source.AddRelationship(chart.GetPackagePart().PartName, TargetMode.Internal, R + "/hyperlink").Id);
                    break;
                case "external":
                    Reference(shape).SetAttribute("id", R,
                        source.AddExternalRelationship("https://example.invalid/chart.xml", XSLFRelation.CHART.Relation).Id);
                    break;
                case "wrongMime":
                    var part = Part(presentation, "/ppt/charts/not-a-chart.xml", "application/xml",
                        Encoding.UTF8.GetBytes("<root/>"));
                    Reference(shape).SetAttribute("id", R,
                        source.AddRelationship(part.PartName, TargetMode.Internal, XSLFRelation.CHART.Relation).Id);
                    break;
            }
            var before = State(presentation);
            Assert.Throws<POIXMLException>(() => shape.GetChart());
            Assert.Throws<POIXMLException>(() => shape.Name = "Changed");
            Assert.Throws<POIXMLException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void MissingChartTargetsAndFragmentReferencesAreNotGuessed()
        {
            using var presentation = Pie();
            var shape = AllShapes(presentation.GetSlides()[0]).OfType<XSLFChartShape>().Single();
            var part = shape.GetChart().GetPackagePart();
            var relationship = shape.GetSheet().GetPackagePart().AddRelationship(
                new Uri(part.PartName.Name + "#unsupported", UriKind.Relative),
                TargetMode.Internal, XSLFRelation.CHART.Relation);
            Reference(shape).SetAttribute("id", R, relationship.Id);
            var before = State(presentation);
            Assert.Throws<NotSupportedException>(() => shape.GetChart());
            Assert.Throws<NotSupportedException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.That(State(presentation), Is.EqualTo(before));
            Reference(shape).SetAttribute("id", R,
                shape.GetSheet().GetPackagePart().GetRelationshipsByType(XSLFRelation.CHART.Relation).GetRelationship(0).Id);
            presentation.Package.RemovePart(part.PartName);
            before = State(presentation);
            Assert.Throws<ArgumentException>(() => shape.GetChart());
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [TestCase("wrongRoot")]
        [TestCase("missingChart")]
        [TestCase("duplicateChart")]
        [TestCase("missingPlotArea")]
        [TestCase("duplicatePlotArea")]
        [TestCase("malformed")]
        [TestCase("dtd")]
        public void InvalidChartXmlBlocksReuseBeforeAddingShapesOrRelationships(string kind)
        {
            using var presentation = Pie();
            var xml = kind switch
            {
                "wrongRoot" => "<c:unknown xmlns:c=\"" + C + "\"/>",
                "missingChart" => "<c:chartSpace xmlns:c=\"" + C + "\"/>",
                "duplicateChart" => "<c:chartSpace xmlns:c=\"" + C + "\"><c:chart/><c:chart/></c:chartSpace>",
                "missingPlotArea" => "<c:chartSpace xmlns:c=\"" + C + "\"><c:chart/></c:chartSpace>",
                "duplicatePlotArea" => "<c:chartSpace xmlns:c=\"" + C + "\"><c:chart><c:plotArea/><c:plotArea/></c:chart></c:chartSpace>",
                "dtd" => "<!DOCTYPE c:chartSpace SYSTEM \"https://example.invalid/never-fetch.dtd\"><c:chartSpace xmlns:c=\"" + C + "\"/>",
                _ => "<malformed"
            };
            var part = Part(presentation, "/ppt/charts/invalid.xml", XSLFRelation.CHART.ContentType, Encoding.UTF8.GetBytes(xml));
            var chart = presentation.GetCharts().Single(item => item.GetPackagePart() == part);
            var slide = presentation.CreateSlide();
            var before = State(presentation);
            Assert.Throws<POIXMLException>(() => chart.GetXmlSnapshot());
            Assert.Throws<POIXMLException>(() => chart.GetChartSpaceSnapshot());
            Assert.Throws<POIXMLException>(() => slide.CreateChart(chart));
            Assert.That(State(presentation), Is.EqualTo(before));
            Assert.That(chart.Data, Is.EqualTo(Encoding.UTF8.GetBytes(xml)));
        }

        [Test]
        public void ForeignNullAndRemovedChartsAreRejectedBeforeMutation()
        {
            using var presentation = Pie();
            using var other = Pie();
            var slide = presentation.CreateSlide();
            var before = State(presentation);
            Assert.Throws<ArgumentNullException>(() => slide.CreateChart(null));
            Assert.Throws<ArgumentException>(() => slide.CreateChart(other.GetCharts().Single()));
            Assert.That(State(presentation), Is.EqualTo(before));
            var chart = presentation.GetCharts().Single();
            presentation.Package.RemovePart(chart.GetPackagePart().PartName);
            before = State(presentation);
            Assert.Throws<InvalidOperationException>(() => slide.CreateChart(chart));
            Assert.Throws<InvalidOperationException>(() => _ = chart.Data);
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void StaleCachedRelationsAreNotReusedForADifferentChartPart()
        {
            using var presentation = Pie();
            var shape = AllShapes(presentation.GetSlides()[0]).OfType<XSLFChartShape>().Single();
            var source = shape.GetSheet().GetPackagePart();
            var part = Part(presentation, "/ppt/charts/other.xml", XSLFRelation.CHART.ContentType, shape.GetChart().Data);
            var id = Reference(shape).GetAttribute("id", R);
            source.RemoveRelationship(id);
            source.AddRelationship(part.PartName, TargetMode.Internal, XSLFRelation.CHART.Relation, id);
            var chart = presentation.GetCharts().Single(item => item.GetPackagePart() == part);
            var before = State(presentation);
            Assert.Throws<POIXMLException>(() => shape.GetSheet().CreateChart(chart));
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [TestCase("shape")]
        [TestCase("group")]
        [TestCase("container")]
        public void DetachedFramesAndMalformedContainersRejectChartMutation(string kind)
        {
            using var presentation = Pie();
            var slide = presentation.CreateSlide();
            var chart = presentation.GetCharts().Single();
            var shape = slide.CreateChart(chart);
            var group = slide.CreateGroup();
            if (kind == "container")
            {
                var tree = slide.GetXmlObject().cSld["spTree", P];
                tree.RemoveChild(tree["grpSpPr", P]);
                var before = State(presentation);
                Assert.Throws<POIXMLException>(() => slide.CreateChart(chart));
                Assert.That(State(presentation), Is.EqualTo(before));
            }
            else
            {
                var xml = kind == "shape" ? shape.GetXmlObject() : group.GetXmlObject();
                xml.ParentNode.RemoveChild(xml);
                var before = State(presentation);
                if (kind == "shape")
                    Assert.Throws<InvalidOperationException>(() => shape.SetAnchor(SKRect.Empty));
                else
                    Assert.Throws<InvalidOperationException>(() => group.CreateChart(chart));
                Assert.That(State(presentation), Is.EqualTo(before));
            }
        }

        [Test]
        public void OpaqueGraphicPayloadsAndMissingTransformsAreNotSilentlyRewritten()
        {
            using var presentation = Pie();
            var shape = AllShapes(presentation.GetSlides()[0]).OfType<XSLFChartShape>().Single();
            var xml = shape.GetXmlObject();
            xml.RemoveChild(xml["xfrm", P]);
            var before = State(presentation);
            Assert.Throws<POIXMLException>(() => shape.SetAnchor(SKRect.Create(1, 2, 3, 4)));
            Assert.That(State(presentation), Is.EqualTo(before));
            var data = xml["graphic", A]["graphicData", A];
            var unknown = xml.OwnerDocument.CreateElement("unknown", "payload", "urn:npoi:test");
            data.AppendChild(unknown);
            before = State(presentation);
            Assert.Throws<NotSupportedException>(() => shape.GetChart());
            Assert.Throws<NotSupportedException>(() => shape.Name = "Changed");
            Assert.That(State(presentation), Is.EqualTo(before));
            data.RemoveChild(unknown);
            data.SetAttribute("uri", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            Assert.That(shape.GetSheet().GetShapes().Single(item => item.ID == shape.ID), Is.TypeOf<XSLFUnsupportedShape>());
            var original = xml.OuterXml;
            using var reopened = Open(Save(presentation));
            Assert.That(AllShapes(reopened.GetSlides()[0]).Single(item => item.ID == shape.ID).GetXmlObject().OuterXml,
                Is.EqualTo(original));
        }

        [TestCase("chart")]
        [TestCase("reference")]
        [TestCase("graphicData")]
        public void AmbiguousChartStructuresRejectEditsWithoutDroppingContent(string kind)
        {
            using var presentation = Pie();
            var shape = AllShapes(presentation.GetSlides()[0]).OfType<XSLFChartShape>().Single();
            if (kind == "chart")
            {
                var part = shape.GetChart().GetPackagePart();
                var root = shape.GetChart().GetXmlSnapshot();
                root.AppendChild(root["chart", C].CloneNode(true));
                Write(part, Encoding.UTF8.GetBytes(root.OuterXml));
            }
            else
            {
                var reference = Reference(shape);
                var element = kind == "reference" ? reference : reference.ParentNode;
                element.ParentNode.AppendChild(element.CloneNode(true));
            }
            var before = State(presentation);
            Assert.Throws<POIXMLException>(() => shape.SetAnchor(SKRect.Empty));
            Assert.Throws<POIXMLException>(() => shape.Name = "Changed");
            Assert.That(State(presentation), Is.EqualTo(before));
        }

        [Test]
        public void ChartFramesRespectSheetWideIdsAndContainerExtensions()
        {
            using var presentation = Pie();
            var slide = presentation.CreateSlide();
            var group = slide.CreateGroup();
            var text = group.CreateTextBox();
            var tree = slide.GetXmlObject().cSld["spTree", P];
            tree.AppendChild(tree.OwnerDocument.CreateElement("p", "extLst", P));
            var chart = slide.CreateChart(presentation.GetCharts().Single());
            Assert.That(chart.ID, Is.Not.EqualTo(group.ID));
            Assert.That(chart.ID, Is.Not.EqualTo(text.ID));
            Assert.That(tree.LastChild.LocalName, Is.EqualTo("extLst"));
            Assert.That(tree.LastChild.PreviousSibling, Is.SameAs(chart.GetXmlObject()));
        }

        [Test]
        public void InvalidFrameCoordinatesDoNotChangeChartOrShapeBytes()
        {
            using var presentation = Pie();
            var shape = AllShapes(presentation.GetSlides()[0]).OfType<XSLFChartShape>().Single();
            var data = shape.GetChart().Data;
            var before = State(presentation);
            Assert.Throws<ArgumentOutOfRangeException>(() => shape.SetAnchor(SKRect.Create(0, 0, -1, 10)));
            Assert.Throws<ArgumentOutOfRangeException>(() => shape.SetAnchor(SKRect.Create(float.NaN, 0, 1, 1)));
            Assert.That(State(presentation), Is.EqualTo(before));
            Assert.That(shape.GetChart().Data, Is.EqualTo(data));
        }

        [Test]
        public void ReadOnlyAndDisposedPresentationsRejectChartFrameMutation()
        {
            using var original = Pie();
            var chart = original.GetCharts().Single();
            using var package = OPCPackage.Open(new MemoryStream(Save(original)), true);
            using var readOnly = new XMLSlideShow(package);
            var saved = AllShapes(readOnly.GetSlides()[0]).OfType<XSLFChartShape>().Single();
            Assert.That(saved.GetChart().GetChartSpaceSnapshot().chart.plotArea.pieChart.Count, Is.EqualTo(1));
            var before = State(readOnly);
            Assert.Throws<InvalidOperationException>(() => saved.Name = "Changed");
            Assert.Throws<InvalidOperationException>(() => saved.SetAnchor(SKRect.Empty));
            Assert.Throws<InvalidOperationException>(() => readOnly.GetSlides()[0].CreateChart(readOnly.GetCharts().Single()));
            Assert.That(State(readOnly), Is.EqualTo(before));
            original.Dispose();
            Assert.Throws<ObjectDisposedException>(() => original.GetCharts());
            Assert.Throws<ObjectDisposedException>(() => _ = chart.Data);
            Assert.Throws<ObjectDisposedException>(() => chart.GetXmlSnapshot());
        }

        [TestCase("pie-chart.pptx")]
        [TestCase("aascu.org_workarea_downloadasset.aspx_id=5864.pptx")]
        public void RealChartPartsAndTheirDependenciesStayBytePreservedAcrossRepeatedSaves(string file)
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(file);
            using var presentation = new XMLSlideShow(stream);
            var charts = presentation.GetCharts();
            Assert.That(charts.Count, Is.GreaterThan(0));
            var dependencies = Dependencies(presentation);
            var frame = presentation.GetSlides().SelectMany(AllShapes).OfType<XSLFChartShape>().First();
            frame.Name = "Edited chart frame only";
            frame.SetAnchor(SKRect.Create(30, 40, 300, 200));
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            Assert.That(again.GetCharts().Count, Is.EqualTo(charts.Count));
            foreach (var dependency in dependencies)
            {
                var part = again.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key));
                Assert.That(Read(part), Is.EqualTo(dependency.Value.Bytes), dependency.Key);
                Assert.That(Relationships(part), Is.EqualTo(dependency.Value.Relationships), dependency.Key);
            }
        }

        private static XMLSlideShow Pie()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            return new XMLSlideShow(input);
        }

        private static IXSLFShapeContainer Container(XMLSlideShow presentation, string kind)
        {
            if (kind == "master")
                return presentation.GetSlideMasters()[0];
            if (kind == "layout")
                return presentation.GetSlideMasters()[0].GetSlideLayouts()[0];
            var slide = presentation.GetSlides().Count == 1 ? presentation.CreateSlide() : presentation.GetSlides().Last();
            if (kind == "slide")
                return slide;
            var group = slide.GetShapes().OfType<XSLFGroupShape>().FirstOrDefault() ?? slide.CreateGroup();
            return kind == "nested"
                ? group.GetShapes().OfType<XSLFGroupShape>().FirstOrDefault() ?? group.CreateGroup() : group;
        }

        private static XmlElement Reference(XSLFChartShape shape)
        {
            return shape.GetXmlObject()["graphic", A]["graphicData", A]["chart", C];
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

        private static Dictionary<string, (byte[] Bytes, string[] Relationships)> Dependencies(XMLSlideShow presentation)
        {
            var slides = new HashSet<PackagePart>(presentation.GetSlides().Select(slide => slide.GetPackagePart()));
            return presentation.Package.GetParts().Where(part => !part.IsRelationshipPart &&
                !slides.Contains(part) && part.ContentType != XSLFRelation.MAIN.ContentType &&
                !part.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(part => part.PartName.Name, part => (Read(part), Relationships(part)));
        }

        private static string[] Relationships(PackagePart part)
        {
            var result = new List<string>();
            foreach (var relationship in part.Relationships)
                result.Add(relationship.Id + "|" + relationship.RelationshipType + "|" +
                    relationship.TargetMode + "|" + relationship.TargetUri.OriginalString);
            return result.ToArray();
        }

        private static string State(XMLSlideShow presentation)
        {
            return string.Join("\n", presentation.GetSlides().Select(slide => slide.GetXmlObject().cSld.OuterXml)) +
                string.Join("\n", presentation.Package.GetParts().Where(part => !part.IsRelationshipPart)
                    .Select(part => part.PartName.Name + "|" + string.Join(",", Relationships(part))));
        }

        private static PackagePart Part(XMLSlideShow presentation, string name, string mime, byte[] bytes)
        {
            var part = presentation.Package.CreatePart(PackagingUriHelper.CreatePartName(name), mime);
            Write(part, bytes);
            return part;
        }

        private static void Write(PackagePart part, byte[] bytes)
        {
            part.Clear();
            using var output = part.GetOutputStream();
            output.Write(bytes, 0, bytes.Length);
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
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XDDF.UserModel.Text;
using NPOI.XSLF.UserModel;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFTables
    {
        private const string P = CT_Presentation.Namespace;
        private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string R = CT_Presentation.RelationshipsNamespace;
        private const long MaxDimension = 27273042316900L;

        [TestCase("slide")]
        [TestCase("layout")]
        [TestCase("master")]
        [TestCase("group")]
        [TestCase("nested")]
        public void AllContainersCreateFixedTablesWithEmptyGeometryAndPreservedText(string kind)
        {
            using var presentation = new XMLSlideShow();
            var container = Container(presentation, kind);
            var snapshot = container.GetShapes();
            var table = container.CreateTable(2, 3);
            Assert.That(table.Parent, Is.SameAs(container));
            Assert.That(snapshot, Is.Empty);
            Assert.That(table.NumberOfRows, Is.EqualTo(2));
            Assert.That(table.NumberOfColumns, Is.EqualTo(3));
            Assert.That(table.GetAnchor(), Is.EqualTo(SKRect.Empty));
            Assert.That(table.GetColumnWidth(0), Is.Zero);
            Assert.That(table.GetRowHeight(0), Is.Zero);
            Assert.That(table.GetCell(0, 0).Text, Is.Empty);
            Assert.That(table.GetCell(0, 0).TextParagraphs.Count, Is.EqualTo(1));
            table.SetAnchor(SKRect.Create(-12, 24, 300, 120));
            Assert.That(table.GetColumnWidth(0), Is.Zero);
            Assert.That(table.GetRowHeight(0), Is.Zero);
            for (var column = 0; column < 3; column++)
                table.SetColumnWidth(column, 70 + column);
            table.SetRowHeight(0, 24);
            table.GetRows()[1].Height = 30;
            table.Name = "Table <&>";
            table.GetCell(0, 0).SetText("  \u4e2d\u6587 \uD83D\uDE03 \t\r\nsecond\r\n");
            var run = table.GetCell(1, 2).TextParagraphs[0].AddNewTextRun();
            run.Text = "Last cell";
            run.FontSize = 18;
            run.IsBold = true;
            Assert.That(table.GetAnchor(), Is.EqualTo(SKRect.Create(-12, 24, 300, 120)));
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            var saved = (XSLFTable)Container(again, kind).GetShapes().Single();
            Assert.That(saved.ID, Is.EqualTo(table.ID));
            Assert.That(saved.Name, Is.EqualTo("Table <&>"));
            Assert.That(saved.GetAnchor(), Is.EqualTo(table.GetAnchor()));
            Assert.That(saved.GetColumnWidth(2), Is.EqualTo(72));
            Assert.That(saved.GetRowHeight(1), Is.EqualTo(30));
            Assert.That(saved.GetCell(0, 0).Text, Is.EqualTo("  \u4e2d\u6587 \uD83D\uDE03 \t\nsecond\n"));
            var savedRun = saved.GetCell(1, 2).TextParagraphs[0].TextRuns[0];
            Assert.That(savedRun.FontSize, Is.EqualTo(18));
            Assert.That(savedRun.IsBold, Is.True);
        }

        [Test]
        public void RowsCellsAndParagraphsHaveCorrectParentsAndReadOnlySnapshots()
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var row = table.GetRows()[0];
            var cell = row.GetCells()[1];
            Assert.That(row.Table, Is.SameAs(table));
            Assert.That(cell.ParentRow, Is.SameAs(row));
            Assert.That(cell.Table, Is.SameAs(table));
            Assert.That(cell.GetSheet(), Is.SameAs(table.GetSheet()));
            Assert.That(((ICollection<XSLFTableRow>)table.GetRows()).IsReadOnly, Is.True);
            Assert.That(((ICollection<XSLFTableCell>)row.GetCells()).IsReadOnly, Is.True);
            var paragraph = cell.TextParagraphs[0];
            Assert.That(paragraph.ParentCell, Is.SameAs(cell));
            Assert.That(paragraph.ParentShape, Is.Null);
            var snapshot = cell.TextParagraphs;
            cell.AddNewTextParagraph().AddNewTextRun().Text = "Added";
            Assert.That(snapshot.Count, Is.EqualTo(1));
            Assert.That(cell.Text, Is.EqualTo("\nAdded"));
            var shape = table.GetSheet().CreateTextBox();
            Assert.That(shape.TextParagraphs[0].ParentShape, Is.SameAs(shape));
            Assert.That(shape.TextParagraphs[0].ParentCell, Is.Null);
        }

        [Test]
        public void TablesShareSheetWideIdsAndAreInsertedBeforeContainerExtensions()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var tree = slide.GetXmlObject().cSld["spTree", P];
            Append(tree, "<p:extLst/>");
            var simple = slide.CreateAutoShape();
            var group = slide.CreateGroup();
            var first = group.CreateTable(1, 1);
            var second = slide.CreateTable(1, 1);
            var text = group.CreateTextBox();
            Assert.That(new[] { simple.ID, group.ID, first.ID, second.ID, text.ID }.Distinct().Count(), Is.EqualTo(5));
            Assert.That(tree.LastChild.LocalName, Is.EqualTo("extLst"));
            Assert.That(tree.LastChild.PreviousSibling.LocalName, Is.EqualTo("graphicFrame"));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetShapes().OfType<XSLFTable>().Single().ID, Is.EqualTo(second.ID));
        }

        [TestCase(0, 1)]
        [TestCase(1, 0)]
        [TestCase(-1, 2)]
        [TestCase(2, -1)]
        public void InvalidGridCountsFailBeforeChangingTheContainer(int rows, int columns)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var before = slide.GetXmlObject().cSld.OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => slide.CreateTable(rows, columns));
            Assert.That(slide.GetXmlObject().cSld.OuterXml, Is.EqualTo(before));
        }

        [TestCase(-1)]
        [TestCase(2)]
        public void InvalidIndexesDoNotChangeTheTable(int index)
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var before = table.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(index, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(0, index));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.GetColumnWidth(index));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetColumnWidth(index, 10));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.GetRowHeight(index));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetRowHeight(index, 10));
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [TestCase(0L)]
        [TestCase(1L)]
        [TestCase(2147483648L)]
        [TestCase(MaxDimension)]
        public void GridDimensionsUseTheirOwnCoordinateBoundsNotFrameExtentBounds(long emus)
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var points = emus / 12700.0;
            table.SetColumnWidth(0, points);
            table.SetRowHeight(0, points);
            Assert.That(TableXml(table)["tblGrid", A]["gridCol", A].GetAttribute("w"), Is.EqualTo(emus.ToString()));
            Assert.That(table.GetRows()[0].GetXmlObject().GetAttribute("h"), Is.EqualTo(emus.ToString()));
            using var reopened = Open(Save(presentation));
            var saved = (XSLFTable)reopened.GetSlides()[0].GetShapes().Single();
            Assert.That(saved.GetColumnWidth(0), Is.EqualTo(points));
            Assert.That(saved.GetRowHeight(0), Is.EqualTo(points));
        }

        [TestCase(double.NaN)]
        [TestCase(double.PositiveInfinity)]
        [TestCase(double.NegativeInfinity)]
        [TestCase(-0.01)]
        [TestCase(double.MaxValue)]
        public void InvalidDimensionsFailBeforeMutation(double value)
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var before = table.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetColumnWidth(0, value));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetRowHeight(0, value));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetColumnWidth(0, (MaxDimension + 1) / 12700.0));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetRowHeight(0, (MaxDimension + 1) / 12700.0));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetAnchor(SKRect.Create(0, 0, -1, 10)));
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void CellFormattingUsesExplicitDefaultsWithoutGuessingPresentationOrTableStyles()
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateTextBox();
            shape.SetText("Shape");
            var table = shape.GetSheet().CreateTable(1, 1);
            Append(TableXml(table)["tblPr", A],
                "<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>");
            var cell = table.GetCell(0, 0);
            var paragraph = cell.TextParagraphs[0];
            var run = paragraph.AddNewTextRun();
            run.Text = "Cell";
            Assert.That(run.FontSize, Is.Null);
            Assert.That(run.IsBold, Is.Null);
            Assert.That(paragraph.TextAlign, Is.Null);
            Append(cell.GetXmlObject()["txBody", A]["lstStyle", A],
                "<a:lvl1pPr algn=\"ctr\"><a:defRPr sz=\"1400\" b=\"1\"><a:latin typeface=\"+mn-lt\"/></a:defRPr></a:lvl1pPr>");
            Assert.That(run.FontSize, Is.EqualTo(14));
            Assert.That(run.IsBold, Is.True);
            Assert.That(run.FontFamily, Is.EqualTo("Arial"));
            Assert.That(paragraph.TextAlign, Is.EqualTo(TextAlignment.Center));
            run.IsBold = false;
            run.IsItalic = true;
            run.IsUnderline = true;
            run.FontSize = 22;
            run.FontFamily = "Courier New";
            paragraph.TextAlign = TextAlignment.Right;
            paragraph.AddLineBreak();
            paragraph.AddNewTextRun().Text = "Second";
            Assert.That(cell.Text, Is.EqualTo("Cell\nSecond"));
            run.IsBold = null;
            Assert.That(run.IsBold, Is.True);
            using var reopened = Open(Save(presentation));
            var saved = reopened.GetSlides()[0].GetShapes().OfType<XSLFTable>().Single().GetCell(0, 0).TextParagraphs[0];
            Assert.That(saved.TextRuns[0].FontSize, Is.EqualTo(22));
            Assert.That(saved.TextRuns[0].IsItalic, Is.True);
            Assert.That(saved.TextRuns[0].IsUnderline, Is.True);
            Assert.That(saved.TextRuns[0].FontFamily, Is.EqualTo("Courier New"));
            Assert.That(saved.TextAlign, Is.EqualTo(TextAlignment.Right));
        }

        [TestCase("")]
        [TestCase("  \t  ")]
        [TestCase("First\r\nsecond\rthird\n")]
        public void WholeCellTextReplacementPreservesWhitespaceAndDetachesOldParagraphs(string text)
        {
            using var presentation = new XMLSlideShow();
            var cell = NewTable(presentation).GetCell(0, 0);
            var old = cell.TextParagraphs[0];
            var oldRun = old.AddNewTextRun();
            oldRun.Text = "Before";
            cell.SetText(text);
            var expected = text.Replace("\r\n", "\n").Replace('\r', '\n');
            Assert.That(cell.Text, Is.EqualTo(expected));
            Assert.Throws<InvalidOperationException>(() => oldRun.Text = "Detached");
            Assert.Throws<InvalidOperationException>(() => old.AddNewTextRun());
            using var reopened = Open(Save(presentation));
            Assert.That(((XSLFTable)reopened.GetSlides()[0].GetShapes().Single()).GetCell(0, 0).Text,
                Is.EqualTo(expected));
        }

        [Test]
        public void CellTextEditsPreserveCellPropertiesStylesAndUnknownRunFormatting()
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var cell = table.GetCell(0, 0);
            var properties = cell.GetXmlObject()["tcPr", A];
            Append(properties, "<a:lnL w=\"12700\"><a:solidFill><a:srgbClr val=\"ABCDEF\"/></a:solidFill></a:lnL>");
            Append(properties, "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>");
            var originalProperties = properties.OuterXml;
            var tableProperties = TableXml(table)["tblPr", A];
            tableProperties.SetAttribute("bandRow", "1");
            Append(tableProperties, "<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>");
            var originalTableProperties = tableProperties.OuterXml;
            cell.SetText("Before");
            var run = cell.TextParagraphs[0].TextRuns[0];
            run.IsBold = true;
            Append(run.GetXmlObject()["rPr", A], "<a:ea typeface=\"Preserved EA\"/>");
            run.Text = "After";
            run.FontSize = 20;
            var before = table.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => cell.SetText("Lossy"));
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
            using var reopened = Open(Save(presentation));
            var savedTable = (XSLFTable)reopened.GetSlides()[0].GetShapes().Single();
            var saved = savedTable.GetCell(0, 0);
            Assert.That(saved.Text, Is.EqualTo("After"));
            Assert.That(saved.GetXmlObject()["tcPr", A].OuterXml, Is.EqualTo(originalProperties));
            Assert.That(TableXml(savedTable)["tblPr", A].OuterXml, Is.EqualTo(originalTableProperties));
            Assert.That(saved.TextParagraphs[0].TextRuns[0].GetXmlObject()["rPr", A]["ea", A].GetAttribute("typeface"),
                Is.EqualTo("Preserved EA"));
        }

        [Test]
        public void FieldsAndHyperlinksArePreservedAndBlockLossyReplacement()
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var cell = table.GetCell(0, 0);
            var paragraph = cell.TextParagraphs[0];
            var run = paragraph.AddNewTextRun();
            run.Text = "Linked ";
            var relationship = table.GetSheet().GetPackagePart().AddExternalRelationship("https://example.invalid/", R + "/hyperlink");
            Append(run.GetXmlObject(), "<a:rPr><a:hlinkClick r:id=\"" + relationship.Id + "\"/></a:rPr>", true);
            Append(paragraph.GetXmlObject(),
                "<a:fld id=\"{7FE94C90-1F39-492A-89D5-F50D2C821A12}\" type=\"datetime1\"><a:t>Cached field</a:t></a:fld>");
            Assert.That(cell.Text, Is.EqualTo("Linked Cached field"));
            var before = table.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => cell.SetText("Lossy"));
            Assert.Throws<NotSupportedException>(() => paragraph.TextRuns.Last().Text = "Field");
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
            run.Text = "Updated ";
            using var reopened = Open(Save(presentation));
            var saved = ((XSLFTable)reopened.GetSlides()[0].GetShapes().Single()).GetCell(0, 0);
            Assert.That(saved.Text, Is.EqualTo("Updated Cached field"));
            Assert.That(saved.TextParagraphs[0].TextRuns[0].GetXmlObject()["rPr", A]["hlinkClick", A]
                .GetAttribute("id", R), Is.EqualTo(relationship.Id));
        }

        [TestCase("rowSpan", "2")]
        [TestCase("gridSpan", "2")]
        [TestCase("hMerge", "true")]
        [TestCase("vMerge", "1")]
        public void MergedTablesAreReadableButAllHighLevelMutationIsRejected(string attribute, string value)
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            table.GetCell(0, 0).SetText("Anchor");
            var cell = table.GetCell(0, 0);
            cell.GetXmlObject().SetAttribute(attribute, value);
            Assert.That(cell.IsMerged, Is.True);
            Assert.That(cell.Text, Is.EqualTo("Anchor"));
            var before = table.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => cell.SetText("Changed"));
            Assert.Throws<NotSupportedException>(() => cell.TextParagraphs[0].TextRuns[0].IsBold = true);
            Assert.Throws<NotSupportedException>(() => table.GetCell(1, 1).AddNewTextParagraph());
            Assert.Throws<NotSupportedException>(() => table.SetColumnWidth(0, 20));
            Assert.Throws<NotSupportedException>(() => table.SetRowHeight(0, 20));
            Assert.Throws<NotSupportedException>(() => table.SetAnchor(SKRect.Empty));
            Assert.Throws<NotSupportedException>(() => table.Name = "Changed");
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
            using var reopened = Open(Save(presentation));
            var saved = (XSLFTable)reopened.GetSlides()[0].GetShapes().Single();
            Assert.That(saved.GetXmlObject().OuterXml, Is.EqualTo(before));
            Assert.That(saved.GetCell(0, 0).IsMerged, Is.True);
        }

        [TestCase("rowSpan", "0")]
        [TestCase("gridSpan", "-1")]
        [TestCase("rowSpan", "2147483648")]
        [TestCase("hMerge", "maybe")]
        public void InvalidMergeMetadataIsNotSilentlyInterpreted(string attribute, string value)
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var cell = table.GetCell(0, 0);
            cell.GetXmlObject().SetAttribute(attribute, value);
            var before = table.GetXmlObject().OuterXml;
            Assert.Throws<POIXMLException>(() => _ = cell.IsMerged);
            Assert.Throws<POIXMLException>(() => cell.SetText("Changed"));
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [TestCase("row")]
        [TestCase("cell")]
        [TestCase("table")]
        public void DetachedWrappersCannotChangeTheTable(string kind)
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var row = table.GetRows()[0];
            var cell = row.GetCells()[0];
            XmlElement node = kind == "row" ? row.GetXmlObject() : kind == "cell" ? cell.GetXmlObject() : table.GetXmlObject();
            node.ParentNode.ReplaceChild(node.CloneNode(true), node);
            var before = ((XSLFSlide)table.GetSheet()).GetXmlObject().cSld.OuterXml;
            if (kind == "row")
                Assert.Throws<InvalidOperationException>(() => row.Height = 10);
            else if (kind == "cell")
                Assert.Throws<InvalidOperationException>(() => cell.SetText("Detached"));
            else
                Assert.Throws<InvalidOperationException>(() => table.SetColumnWidth(0, 10));
            Assert.That(((XSLFSlide)table.GetSheet()).GetXmlObject().cSld.OuterXml, Is.EqualTo(before));
        }

        [TestCase("missingGrid")]
        [TestCase("duplicateGrid")]
        [TestCase("missingBody")]
        [TestCase("duplicateBody")]
        [TestCase("invalidWidth")]
        [TestCase("invalidHeight")]
        public void MalformedTableStructureFailsExplicitlyBeforeMutation(string kind)
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var xml = TableXml(table);
            var cell = table.GetCell(0, 0);
            switch (kind)
            {
                case "missingGrid": xml.RemoveChild(xml["tblGrid", A]); break;
                case "duplicateGrid": xml.AppendChild(xml["tblGrid", A].CloneNode(true)); break;
                case "missingBody": cell.GetXmlObject().RemoveChild(cell.GetXmlObject()["txBody", A]); break;
                case "duplicateBody": cell.GetXmlObject().AppendChild(cell.GetXmlObject()["txBody", A].CloneNode(true)); break;
                case "invalidWidth": xml["tblGrid", A]["gridCol", A].SetAttribute("w", "-1"); break;
                case "invalidHeight": table.GetRows()[0].GetXmlObject().SetAttribute("h", "NaN"); break;
            }
            var before = table.GetXmlObject().OuterXml;
            Assert.Throws<POIXMLException>(() => table.SetAnchor(SKRect.Empty));
            Assert.Throws<POIXMLException>(() => cell.SetText("Changed"));
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void UnknownAndIrregularStructuresArePreservedWithoutInventingGridPositions()
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var row = table.GetRows()[0];
            row.GetXmlObject().RemoveChild(row.GetCells()[0].GetXmlObject());
            var before = table.GetXmlObject().OuterXml;
            Assert.That(row.GetCells().Count, Is.EqualTo(1));
            Assert.Throws<NotSupportedException>(() => table.SetColumnWidth(0, 10));
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
            var other = table.GetRows()[1].GetCells()[0];
            Append(other.GetXmlObject(), "<unknown:content xmlns:unknown=\"urn:npoi:test\"/>");
            before = table.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => _ = other.Text);
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(before));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetShapes().Single().GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void UnknownGraphicFramesStayOpaqueAndMalformedFrameTransformsAreNotRecreated()
        {
            using var presentation = new XMLSlideShow();
            var table = NewTable(presentation);
            var xml = table.GetXmlObject();
            var transform = xml["xfrm", P];
            xml.RemoveChild(transform);
            var before = xml.OuterXml;
            Assert.Throws<POIXMLException>(() => table.SetAnchor(SKRect.Create(1, 2, 3, 4)));
            Assert.That(xml.OuterXml, Is.EqualTo(before));
            xml["graphic", A]["graphicData", A].SetAttribute("uri", "urn:npoi:unknown");
            Assert.That(table.GetSheet().GetShapes().Single(), Is.TypeOf<XSLFUnsupportedShape>());
            before = xml.OuterXml;
            Assert.Throws<NotSupportedException>(() => table.SetColumnWidth(0, 10));
            Assert.That(xml.OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void ReadOnlyDisposedAndRemovedSheetsRejectTableChanges()
        {
            using var original = new XMLSlideShow();
            var table = NewTable(original);
            using var package = OPCPackage.Open(new MemoryStream(Save(original)), true);
            using var readOnly = new XMLSlideShow(package);
            var saved = (XSLFTable)readOnly.GetSlides()[0].GetShapes().Single();
            Assert.That(saved.GetCell(0, 0).Text, Is.Empty);
            var before = saved.GetXmlObject().OuterXml;
            Assert.Throws<InvalidOperationException>(() => saved.GetCell(0, 0).SetText("Changed"));
            Assert.Throws<InvalidOperationException>(() => saved.SetRowHeight(0, 20));
            Assert.Throws<InvalidOperationException>(() => readOnly.GetSlides()[0].CreateTable(1, 1));
            Assert.That(saved.GetXmlObject().OuterXml, Is.EqualTo(before));
            original.RemoveSlide(0);
            Assert.Throws<InvalidOperationException>(() => table.SetColumnWidth(0, 10));
            original.Dispose();
            Assert.Throws<ObjectDisposedException>(() => table.SetRowHeight(0, 20));
        }

        [TestCase("sample.pptx")]
        [TestCase("shapes.pptx")]
        public void AddingTablesToRealFilesPreservesExistingShapesAndDependencyBytes(string file)
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(file);
            using var presentation = new XMLSlideShow(input);
            var slide = presentation.GetSlides()[0];
            var originals = slide.GetShapes().Select(shape => shape.GetXmlObject().OuterXml).ToArray();
            var dependencies = presentation.Package.GetParts().Where(part => !part.IsRelationshipPart &&
                part != slide.GetPackagePart() && part.ContentType != XSLFRelation.MAIN.ContentType &&
                !part.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(part => part.PartName.Name, Read);
            var table = slide.CreateTable(2, 2);
            table.SetColumnWidth(0, 120);
            table.SetRowHeight(0, 24);
            table.GetCell(0, 0).SetText("Inserted \u4e2d\u6587");
            table.SetAnchor(SKRect.Create(30, 40, 240, 48));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetShapes().Take(originals.Length)
                .Select(shape => shape.GetXmlObject().OuterXml), Is.EqualTo(originals));
            foreach (var dependency in dependencies)
                Assert.That(Read(reopened.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key))),
                    Is.EqualTo(dependency.Value), dependency.Key);
        }

        [TestCase("shapes.pptx")]
        [TestCase("present1.pptx")]
        [TestCase("with_japanese.pptx")]
        public void NativeUnmergedTableTextCanBeEditedWithoutChangingOtherCellsOrDependencies(string file)
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(file);
            using var presentation = new XMLSlideShow(input);
            var table = presentation.GetSlides().SelectMany(AllShapes).OfType<XSLFTable>()
                .First(item => item.GetRows().SelectMany(row => row.GetCells()).All(cell => !cell.IsMerged));
            var run = table.GetRows().SelectMany(row => row.GetCells())
                .SelectMany(cell => cell.TextParagraphs).SelectMany(paragraph => paragraph.TextRuns)
                .First(item => item.GetXmlObject().LocalName == "r");
            var cell = run.ParentParagraph.ParentCell;
            var otherCells = table.GetRows().SelectMany(row => row.GetCells())
                .Where(item => !ReferenceEquals(item.GetXmlObject(), cell.GetXmlObject()))
                .Select(item => item.GetXmlObject().OuterXml).ToArray();
            var properties = cell.GetXmlObject()["tcPr", A]?.OuterXml;
            var sheet = table.GetSheet();
            var dependencies = presentation.Package.GetParts().Where(part => !part.IsRelationshipPart &&
                part != sheet.GetPackagePart() && part.ContentType != XSLFRelation.MAIN.ContentType &&
                !part.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(part => part.PartName.Name, Read);
            run.Text = "Edited native table \u4e2d\u6587";
            run.FontSize = 22;
            using var reopened = Open(Save(presentation));
            var savedSheet = reopened.GetSlides().Single(item => item.GetPackagePart().PartName.Equals(sheet.GetPackagePart().PartName));
            var saved = AllShapes(savedSheet).OfType<XSLFTable>().Single(item => item.ID == table.ID);
            var changed = saved.GetRows().SelectMany(row => row.GetCells())
                .Single(item => item.Text.Contains("Edited native table"));
            Assert.That(changed.GetXmlObject()["tcPr", A]?.OuterXml, Is.EqualTo(properties));
            Assert.That(saved.GetRows().SelectMany(row => row.GetCells())
                .Where(item => !ReferenceEquals(item.GetXmlObject(), changed.GetXmlObject()))
                .Select(item => item.GetXmlObject().OuterXml), Is.EqualTo(otherCells));
            foreach (var dependency in dependencies)
                Assert.That(Read(reopened.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key))),
                    Is.EqualTo(dependency.Value), dependency.Key);
        }

        [Test]
        public void NativeMergedTablesStayReadableAndBytePreservedWhenEditingAnotherSlide()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("shapes.pptx");
            using var presentation = new XMLSlideShow(input);
            var table = presentation.GetSlides().SelectMany(AllShapes).OfType<XSLFTable>()
                .First(item => item.GetRows().SelectMany(row => row.GetCells()).Any(cell => cell.IsMerged));
            var xml = table.GetXmlObject().OuterXml;
            var part = table.GetSheet().GetPackagePart();
            var originalBytes = Read(part);
            var originalText = table.GetRows().SelectMany(row => row.GetCells()).Select(cell => cell.Text).ToArray();
            Assert.Throws<NotSupportedException>(() => table.GetCell(0, 0).SetText("Changed"));
            Assert.That(table.GetXmlObject().OuterXml, Is.EqualTo(xml));
            presentation.GetSlides().First(slide => slide != table.GetSheet()).CreateTable(1, 1);
            using var reopened = Open(Save(presentation));
            var saved = reopened.GetSlides().Single(slide => slide.GetPackagePart().PartName.Equals(part.PartName));
            Assert.That(Read(saved.GetPackagePart()), Is.EqualTo(originalBytes));
            var savedTable = AllShapes(saved).OfType<XSLFTable>().Single(item => item.ID == table.ID);
            Assert.That(savedTable.GetRows().SelectMany(row => row.GetCells()).Select(cell => cell.Text),
                Is.EqualTo(originalText));
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

        private static XSLFTable NewTable(XMLSlideShow presentation)
        {
            return presentation.CreateSlide().CreateTable(2, 2);
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

        private static XmlElement TableXml(XSLFTable table)
        {
            return table.GetXmlObject()["graphic", A]["graphicData", A]["tbl", A];
        }

        private static void Append(XmlElement parent, string content, bool prepend = false)
        {
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml("<root xmlns:a=\"" + A + "\" xmlns:p=\"" + P + "\" xmlns:r=\"" + R + "\">" + content + "</root>");
            var element = parent.OwnerDocument.ImportNode(document.DocumentElement.FirstChild, true);
            if (prepend)
                parent.PrependChild(element);
            else
                parent.AppendChild(element);
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

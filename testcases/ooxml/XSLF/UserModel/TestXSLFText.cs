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
    public class TestXSLFText
    {
        private const string P = CT_Presentation.Namespace;
        private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        [TestCase("slide")]
        [TestCase("layout")]
        [TestCase("master")]
        [TestCase("group")]
        public void ContainersCreateTypedTextBoxesWithStableIdsAndUnicodeText(string kind)
        {
            using var presentation = new XMLSlideShow();
            IXSLFShapeContainer container = Container(presentation, kind);
            var box = container.CreateTextBox();
            Assert.That(box.Parent, Is.SameAs(container));
            Assert.That(box.TextParagraphs.Count, Is.EqualTo(1));
            Assert.That(box.Text, Is.Empty);
            Assert.That(box.TextParagraphs[0].TextRuns, Is.Empty);
            box.SetAnchor(SKRect.Create(12, 24, 240, 90));
            box.SetText("  Hello <&> \u4e2d\u6587 \uD83D\uDE03  \r\n\tsecond\r\n");
            var id = box.ID;
            var expected = "  Hello <&> \u4e2d\u6587 \uD83D\uDE03  \n\tsecond\n";
            using var reopened = Open(Save(presentation));
            var saved = (XSLFTextBox)Container(reopened, kind).GetShapes().Single();
            Assert.That(saved.ID, Is.EqualTo(id));
            Assert.That(saved.Text, Is.EqualTo(expected));
            Assert.That(saved.TextParagraphs.Count, Is.EqualTo(3));
            Assert.That(saved.GetAnchor(), Is.EqualTo(box.GetAnchor()));
            saved.TextParagraphs[0].TextRuns[0].Text = "Updated \u4e16\u754c";
            using var again = Open(Save(reopened));
            Assert.That(((XSLFSimpleShape)Container(again, kind).GetShapes()[0]).Text,
                Is.EqualTo("Updated \u4e16\u754c\n\tsecond\n"));
        }

        [TestCase("")]
        [TestCase(" ")]
        [TestCase("  \t  ")]
        [TestCase("\n")]
        [TestCase("a\r\nb\rc\n")]
        public void WhitespaceAndParagraphBreaksSurviveRepeatedSaves(string text)
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            Assert.That(shape.TextParagraphs, Is.Empty);
            shape.SetText(text);
            var expected = text.Replace("\r\n", "\n").Replace('\r', '\n');
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            Assert.That(((XSLFSimpleShape)again.GetSlides()[0].GetShapes()[0]).Text, Is.EqualTo(expected));
        }

        [Test]
        public void ParagraphsRunsAndBreaksFollowXmlOrderWithoutMovingEndProperties()
        {
            using var presentation = new XMLSlideShow();
            var shape = presentation.CreateSlide().CreateAutoShape();
            var paragraph = shape.AddNewTextParagraph();
            Append(paragraph.GetXmlObject(), "<a:endParaRPr lang=\"en-US\"/>");
            var snapshot = paragraph.TextRuns;
            var first = paragraph.AddNewTextRun();
            first.Text = "First ";
            paragraph.AddLineBreak();
            paragraph.AddNewTextRun().Text = "second";
            shape.AddNewTextParagraph().AddNewTextRun().Text = "Third";
            Assert.That(snapshot, Is.Empty);
            Assert.That(((ICollection<XSLFTextParagraph>)shape.TextParagraphs).IsReadOnly, Is.True);
            Assert.That(((ICollection<XSLFTextRun>)paragraph.TextRuns).IsReadOnly, Is.True);
            Assert.That(paragraph.GetXmlObject().LastChild.LocalName, Is.EqualTo("endParaRPr"));
            Assert.That(first.ParentParagraph, Is.SameAs(paragraph));
            Assert.That(paragraph.ParentShape, Is.SameAs(shape));
            Assert.That(shape.Text, Is.EqualTo("First \nsecond\nThird"));
            using var reopened = Open(Save(presentation));
            Assert.That(((XSLFSimpleShape)reopened.GetSlides()[0].GetShapes()[0]).Text, Is.EqualTo(shape.Text));
        }

        [TestCase(1d, "100", 1d)]
        [TestCase(4000d, "400000", 4000d)]
        [TestCase(11.125d, "1112", 11.12d)]
        [TestCase(11.129d, "1113", 11.13d)]
        public void FontSizesUseHundredthsOfAPointAndExactBounds(double points, string emSize, double expected)
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            run.FontSize = points;
            Assert.That(run.GetXmlObject()["rPr", A].GetAttribute("sz"), Is.EqualTo(emSize));
            using var reopened = Open(Save(presentation));
            Assert.That(FirstRun(reopened).FontSize, Is.EqualTo(expected));
        }

        [TestCase(double.NaN)]
        [TestCase(double.PositiveInfinity)]
        [TestCase(double.NegativeInfinity)]
        [TestCase(0.99d)]
        [TestCase(4000.01d)]
        public void InvalidFontSizesDoNotCreateOrChangeProperties(double invalid)
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            var before = run.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSize = invalid);
            Assert.That(run.GetXmlObject().OuterXml, Is.EqualTo(before));
            run.IsBold = true;
            before = run.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSize = invalid);
            Assert.That(run.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void DirectFormattingSupportsFalseAndUnsetWithoutInventedDefaults()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            Assert.That(run.FontSize, Is.Null);
            Assert.That(run.FontFamily, Is.Null);
            Assert.That(run.IsBold, Is.Null);
            Assert.That(run.IsItalic, Is.Null);
            Assert.That(run.IsUnderline, Is.Null);
            Assert.That(run.ParentParagraph.TextAlign, Is.Null);
            run.Text = "Rich \u4e2d\u6587";
            run.FontFamily = "Calibri";
            run.FontSize = 18;
            run.IsBold = true;
            run.IsItalic = false;
            run.IsUnderline = true;
            run.ParentParagraph.TextAlign = TextAlignment.Center;
            using var reopened = Open(Save(presentation));
            var saved = FirstRun(reopened);
            Assert.That(saved.FontFamily, Is.EqualTo("Calibri"));
            Assert.That(saved.FontSize, Is.EqualTo(18));
            Assert.That(saved.IsBold, Is.True);
            Assert.That(saved.IsItalic, Is.False);
            Assert.That(saved.IsUnderline, Is.True);
            Assert.That(saved.ParentParagraph.TextAlign, Is.EqualTo(TextAlignment.Center));
            saved.FontSize = null;
            saved.FontFamily = null;
            saved.IsBold = null;
            saved.IsItalic = null;
            saved.IsUnderline = null;
            saved.ParentParagraph.TextAlign = null;
            Assert.That(saved.FontFamily, Is.Null);
            Assert.That(saved.FontSize, Is.Null);
            Assert.That(saved.IsBold, Is.Null);
            Assert.That(saved.IsItalic, Is.Null);
            Assert.That(saved.IsUnderline, Is.Null);
            Assert.That(saved.ParentParagraph.TextAlign, Is.Null);
        }

        [Test]
        public void SupportedPropertiesResolveIndependentlyAcrossPlaceholderAndLocalStyles()
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var masterShape = master.CreateTextBox();
            Placeholder(masterShape, "body", "0");
            ListStyle(masterShape, "<a:lvl1pPr algn=\"r\"><a:defRPr sz=\"2200\" b=\"1\" u=\"sng\">" +
                "<a:latin typeface=\"+mn-lt\"/></a:defRPr></a:lvl1pPr>");
            var layout = master.GetSlideLayouts()[0].CreateTextBox();
            layout.SetText("Layout prompt must not become slide text");
            Placeholder(layout, "obj", "7");
            ListStyle(layout, "<a:lvl1pPr algn=\"ctr\"><a:defRPr i=\"1\"/></a:lvl1pPr>");
            var box = presentation.CreateSlide().CreateTextBox();
            Placeholder(box, "body", "7");
            var run = box.TextParagraphs[0].AddNewTextRun();
            run.Text = "Actual slide text";
            Assert.That(run.FontSize, Is.EqualTo(22));
            Assert.That(run.FontFamily, Is.EqualTo("Arial"));
            Assert.That(run.IsBold, Is.True);
            Assert.That(run.IsItalic, Is.True);
            Assert.That(run.IsUnderline, Is.True);
            Assert.That(run.ParentParagraph.TextAlign, Is.EqualTo(TextAlignment.Center));
            Assert.That(box.Text, Is.EqualTo("Actual slide text"));
            ListStyle(box, "<a:lvl1pPr><a:defRPr sz=\"2500\"/></a:lvl1pPr>");
            Append(run.ParentParagraph.GetXmlObject(), "<a:pPr><a:defRPr sz=\"2700\"/></a:pPr>", true);
            Assert.That(run.FontSize, Is.EqualTo(27));
            run.FontSize = 30;
            Assert.That(run.FontSize, Is.EqualTo(30));
            run.FontSize = null;
            Assert.That(run.FontSize, Is.EqualTo(27));
            run.ParentParagraph.GetXmlObject()["pPr", A]["defRPr", A].RemoveAttribute("sz");
            Assert.That(run.FontSize, Is.EqualTo(25));
            ListStyle(box, "");
            Assert.That(run.FontSize, Is.EqualTo(22));
            run.IsBold = false;
            run.IsUnderline = false;
            Assert.That(run.IsBold, Is.False);
            Assert.That(run.IsUnderline, Is.False);
            run.IsBold = null;
            run.IsUnderline = null;
            Assert.That(run.IsBold, Is.True);
            Assert.That(run.IsUnderline, Is.True);
            using var reopened = Open(Save(presentation));
            Assert.That(FirstRun(reopened).FontSize, Is.EqualTo(22));
            Assert.That(FirstRun(reopened).IsItalic, Is.True);
            Assert.That(((XSLFSimpleShape)reopened.GetSlides()[0].GetShapes()[0]).Text, Is.EqualTo("Actual slide text"));
        }

        [Test]
        public void PresentationDefaultsApplyToPlainTextAndMasterStylesApplyToPlaceholders()
        {
            using var presentation = new XMLSlideShow();
            presentation.GetCTPresentation().AnyElements = new[]
            {
                Xml("<p:defaultTextStyle><a:lvl1pPr algn=\"r\"><a:defRPr sz=\"1800\"/></a:lvl1pPr></p:defaultTextStyle>")
            };
            var master = presentation.GetSlideMasters()[0];
            master.GetXmlObject().AnyElements = new[]
            {
                Xml("<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz=\"3600\"/></a:lvl1pPr></p:titleStyle>" +
                    "<p:bodyStyle><a:lvl1pPr><a:defRPr sz=\"2400\"/></a:lvl1pPr></p:bodyStyle>" +
                    "<p:otherStyle><a:lvl1pPr><a:defRPr b=\"1\"/></a:lvl1pPr></p:otherStyle></p:txStyles>")
            };
            var plain = presentation.CreateSlide().CreateTextBox();
            var run = plain.TextParagraphs[0].AddNewTextRun();
            Assert.That(run.FontSize, Is.EqualTo(18));
            Assert.That(run.IsBold, Is.True);
            Assert.That(run.ParentParagraph.TextAlign, Is.EqualTo(TextAlignment.Right));
            Placeholder(plain, "title", "15");
            Assert.That(run.FontSize, Is.EqualTo(36));
            Assert.That(run.ParentParagraph.TextAlign, Is.Null, "Presentation defaults do not override placeholder defaults.");
            Placeholder(plain, "body", "15");
            Assert.That(run.FontSize, Is.EqualTo(24));
        }

        [Test]
        public void AllNineListLevelsAndDefaultPropertiesAreReadWithoutChangingXml()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            var shape = run.ParentParagraph.ParentShape;
            ListStyle(shape, "<a:defPPr><a:defRPr i=\"1\"/></a:defPPr>" + string.Concat(Enumerable.Range(0, 9)
                .Select(level => "<a:lvl" + (level + 1) + "pPr><a:defRPr sz=\"" + (1200 + level * 100) +
                    "\"/></a:lvl" + (level + 1) + "pPr>")));
            for (var level = 0; level < 9; level++)
            {
                run.ParentParagraph.IndentLevel = level;
                var before = shape.GetXmlObject().OuterXml;
                Assert.That(run.FontSize, Is.EqualTo(12 + level));
                Assert.That(run.IsItalic, Is.True);
                Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            }
            var original = shape.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => run.ParentParagraph.IndentLevel = -1);
            Assert.Throws<ArgumentOutOfRangeException>(() => run.ParentParagraph.IndentLevel = 9);
            Assert.Throws<ArgumentOutOfRangeException>(() => run.ParentParagraph.TextAlign = (TextAlignment)int.MaxValue);
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(original));
        }

        [Test]
        public void AmbiguousPlaceholderMatchesFailUnlessTheRequestedPropertyIsExplicit()
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var layout = master.GetSlideLayouts()[0];
            Placeholder(layout.CreateTextBox(), "body", "2");
            Placeholder(layout.CreateTextBox(), "obj", "2");
            var box = presentation.CreateSlide().CreateTextBox();
            Placeholder(box, "body", "2");
            var run = box.TextParagraphs[0].AddNewTextRun();
            Assert.Throws<POIXMLException>(() => _ = run.FontSize);
            run.FontSize = 28;
            Assert.That(run.FontSize, Is.EqualTo(28));
            Assert.That(box.Text, Is.Empty);
        }

        [TestCase("+mj-lt")]
        [TestCase("+mn-lt")]
        public void LatinThemeFontsResolveAndMissingThemeFailsExplicitly(string token)
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            run.FontFamily = token;
            Assert.That(run.FontFamily, Is.EqualTo("Arial"));
            Assert.That(run.GetXmlObject()["rPr", A]["latin", A].GetAttribute("typeface"), Is.EqualTo(token));
            var master = presentation.GetSlideMasters()[0].GetPackagePart();
            var theme = master.GetRelationshipsByType(XSLFRelation.THEME.Relation).GetRelationship(0);
            master.RemoveRelationship(theme.Id);
            Assert.Throws<POIXMLException>(() => _ = run.FontFamily);
        }

        [TestCase("b", "bad")]
        [TestCase("i", "2")]
        [TestCase("u", "unknown")]
        [TestCase("sz", "99")]
        [TestCase("sz", "400001")]
        [TestCase("sz", "bad")]
        public void MalformedSupportedPropertiesFailWithoutSilentDefaults(string property, string value)
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            run.IsBold = false;
            run.GetXmlObject()["rPr", A].SetAttribute(property, value);
            Assert.Throws<POIXMLException>(() =>
            {
                switch (property)
                {
                    case "b": _ = run.IsBold; break;
                    case "i": _ = run.IsItalic; break;
                    case "u": _ = run.IsUnderline; break;
                    default: _ = run.FontSize; break;
                }
            });
        }

        [Test]
        public void WholeReplacementClearsOnlySupportedFormattingAndDetachesOldWrappers()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            var paragraph = run.ParentParagraph;
            var shape = paragraph.ParentShape;
            run.FontSize = 20;
            run.IsBold = true;
            run.FontFamily = "Calibri";
            paragraph.TextAlign = TextAlignment.Center;
            var bodyPr = shape.GetXmlObject()["txBody", P]["bodyPr", A];
            bodyPr.SetAttribute("wrap", "none");
            shape.SetText("Replacement\nNext");
            Assert.That(shape.Text, Is.EqualTo("Replacement\nNext"));
            Assert.That(shape.TextParagraphs[0].TextRuns[0].FontSize, Is.Null);
            Assert.That(bodyPr.GetAttribute("wrap"), Is.EqualTo("none"));
            Assert.Throws<InvalidOperationException>(() => run.Text = "Detached");
            Assert.Throws<InvalidOperationException>(() => paragraph.AddNewTextRun());
        }

        [TestCase("field")]
        [TestCase("hyperlink")]
        [TestCase("unknown")]
        [TestCase("format")]
        public void UnsupportedContentBlocksWholeReplacementWithoutChangingAnyXml(string kind)
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            run.Text = "Original";
            var paragraph = run.ParentParagraph;
            if (kind == "field")
                Append(paragraph.GetXmlObject(), "<a:fld id=\"{11111111-1111-1111-1111-111111111111}\" type=\"slidenum\">" +
                    "<a:rPr/><a:t>1</a:t></a:fld>");
            else if (kind == "unknown")
                Append(paragraph.GetXmlObject(), "<test:opaque xmlns:test=\"urn:npoi:test\">opaque</test:opaque>");
            else
            {
                run.IsBold = true;
                Append(run.GetXmlObject()["rPr", A], kind == "hyperlink" ? "<a:hlinkClick action=\"ppaction://nextslide\"/>" :
                    "<a:solidFill><a:srgbClr val=\"ABCDEF\"/></a:solidFill>");
            }
            var shape = paragraph.ParentShape;
            var unsupported = kind == "field" || kind == "unknown" ? paragraph.GetXmlObject().LastChild :
                run.GetXmlObject()["rPr", A].LastChild;
            var unsupportedXml = unsupported.OuterXml;
            var before = shape.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => shape.SetText("Must not replace"));
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            run.Text = "Edited regular run";
            run.FontSize = 22;
            if (kind == "field")
            {
                var field = paragraph.TextRuns.Last();
                Assert.That(field.Text, Is.EqualTo("1"));
                Assert.Throws<NotSupportedException>(() => field.Text = "2");
                Assert.Throws<NotSupportedException>(() => field.FontSize = 12);
            }
            else if (kind == "unknown")
                Assert.Throws<NotSupportedException>(() => _ = shape.Text);
            using var reopened = Open(Save(presentation));
            var saved = (XSLFSimpleShape)reopened.GetSlides()[0].GetShapes()[0];
            Assert.That(saved.TextParagraphs[0].TextRuns[0].Text, Is.EqualTo("Edited regular run"));
            var savedUnsupported = kind == "field" || kind == "unknown"
                ? saved.TextParagraphs[0].GetXmlObject().LastChild
                : saved.TextParagraphs[0].TextRuns[0].GetXmlObject()["rPr", A].LastChild;
            Assert.That(savedUnsupported.OuterXml, Is.EqualTo(unsupportedXml));
            if (kind == "field")
                Assert.That(saved.TextParagraphs[0].TextRuns.Last().Text, Is.EqualTo("1"));
        }

        [Test]
        public void InvalidTextAndFontsFailBeforeMutationAndRunsRequireExplicitBreaks()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            var shape = run.ParentParagraph.ParentShape;
            var before = shape.GetXmlObject().OuterXml;
            Assert.Throws<ArgumentNullException>(() => shape.SetText(null));
            Assert.Throws<XmlException>(() => shape.SetText("\u0001"));
            Assert.Throws<ArgumentNullException>(() => run.Text = null);
            Assert.Throws<XmlException>(() => run.Text = "\uD800");
            Assert.Throws<ArgumentException>(() => run.Text = "a\nb");
            Assert.Throws<ArgumentException>(() => run.FontFamily = " ");
            Assert.Throws<NotSupportedException>(() => run.FontFamily = "+mn-ea");
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            var lineBreak = run.ParentParagraph.AddLineBreak();
            Assert.Throws<NotSupportedException>(() => lineBreak.Text = "Not a break");
            lineBreak.IsBold = true;
            Assert.That(lineBreak.IsBold, Is.True);
        }

        [Test]
        public void HyperlinkRelationshipsAndUnsupportedRunPropertiesSurviveTextAndFontEdits()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            run.Text = "Linked text";
            run.IsBold = true;
            var part = run.ParentParagraph.ParentShape.GetSheet().GetPackagePart();
            var relationship = part.AddExternalRelationship("https://example.org/",
                CT_Presentation.RelationshipsNamespace + "/hyperlink");
            var properties = run.GetXmlObject()["rPr", A];
            Append(properties, "<a:ea typeface=\"Existing EA\"/>");
            var link = properties.OwnerDocument.CreateElement("a", "hlinkClick", A);
            var id = properties.OwnerDocument.CreateAttribute("r", "id", CT_Presentation.RelationshipsNamespace);
            id.Value = relationship.Id;
            link.SetAttributeNode(id);
            properties.AppendChild(link);
            var originalLink = link.OuterXml;
            run.FontFamily = "Calibri";
            run.FontSize = 26;
            run.Text = "Updated linked text";
            Assert.That(properties.ChildNodes.Cast<XmlNode>().Select(node => node.LocalName),
                Is.EqualTo(new[] { "latin", "ea", "hlinkClick" }));
            Assert.That(link.OuterXml, Is.EqualTo(originalLink));
            using var reopened = Open(Save(presentation));
            var saved = FirstRun(reopened);
            Assert.That(saved.Text, Is.EqualTo("Updated linked text"));
            Assert.That(saved.GetXmlObject()["rPr", A]["hlinkClick", A].OuterXml, Is.EqualTo(originalLink));
            Assert.That(saved.GetXmlObject()["rPr", A]["ea", A].GetAttribute("typeface"), Is.EqualTo("Existing EA"));
            var preserved = reopened.GetSlides()[0].GetPackagePart().GetRelationship(relationship.Id);
            Assert.That(preserved.TargetMode, Is.EqualTo(TargetMode.External));
            Assert.That(preserved.TargetUri.OriginalString, Is.EqualTo("https://example.org/"));
        }

        [Test]
        public void OpaqueTextBodyNestedTextAndDuplicatePropertiesAreNotSilentlyRewritten()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            var shape = run.ParentParagraph.ParentShape;
            var text = run.GetXmlObject()["t", A];
            Append(text, "<test:unknown xmlns:test=\"urn:npoi:test\"/>");
            var before = shape.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => _ = run.Text);
            Assert.Throws<NotSupportedException>(() => run.Text = "New");
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            text.RemoveAll();
            run.IsBold = true;
            run.GetXmlObject().AppendChild(run.GetXmlObject()["rPr", A].CloneNode(true));
            before = shape.GetXmlObject().OuterXml;
            Assert.Throws<POIXMLException>(() => run.FontSize = 20);
            Assert.Throws<POIXMLException>(() => shape.SetText("New"));
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
            run.GetXmlObject().RemoveChild(run.GetXmlObject().LastChild);
            Append(shape.GetXmlObject()["txBody", P], "<test:opaque xmlns:test=\"urn:npoi:test\"/>");
            before = shape.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => _ = shape.Text);
            Assert.Throws<NotSupportedException>(() => shape.SetText("New"));
            Assert.Throws<NotSupportedException>(() => shape.AddNewTextParagraph());
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void MissingListLevelsDoNotInventValuesAndMalformedLevelsAndAlignmentFail()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            ListStyle(run.ParentParagraph.ParentShape, "<a:lvl1pPr><a:defRPr sz=\"2000\"/></a:lvl1pPr>");
            run.ParentParagraph.IndentLevel = 1;
            Assert.That(run.FontSize, Is.Null);
            var properties = run.ParentParagraph.GetXmlObject()["pPr", A];
            properties.SetAttribute("lvl", "9");
            Assert.Throws<POIXMLException>(() => _ = run.FontSize);
            properties.SetAttribute("lvl", "0");
            properties.SetAttribute("algn", "bad");
            Assert.Throws<POIXMLException>(() => _ = run.ParentParagraph.TextAlign);
        }

        [TestCase("index")]
        [TestCase("type")]
        [TestCase("master")]
        [TestCase("opaque")]
        public void InvalidOrOpaquePlaceholderMetadataDoesNotProduceGuessedFormatting(string kind)
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            var shape = run.ParentParagraph.ParentShape;
            Placeholder(shape, kind == "type" ? "invalid" : "body", kind == "index" ? "invalid" : "5");
            var master = presentation.GetSlideMasters()[0];
            if (kind == "master")
            {
                Placeholder(master.CreateTextBox(), "body", "0");
                Placeholder(master.CreateTextBox(), "body", "1");
            }
            if (kind == "opaque")
            {
                var layout = master.GetSlideLayouts()[0];
                Append(layout.GetXmlObject().cSld["spTree", P], "<test:opaque xmlns:test=\"urn:npoi:test\">" +
                    "<p:ph type=\"body\" idx=\"5\"/></test:opaque>");
                Assert.Throws<NotSupportedException>(() => _ = run.FontSize);
            }
            else
                Assert.Throws<POIXMLException>(() => _ = run.FontSize);
        }

        [Test]
        public void UnknownContentInsideRegularRunsBlocksInterpretationAndEditsWithoutLoss()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            Append(run.GetXmlObject(), "<test:opaque xmlns:test=\"urn:npoi:test\"/>");
            var before = run.GetXmlObject().OuterXml;
            Assert.Throws<NotSupportedException>(() => _ = run.Text);
            Assert.Throws<NotSupportedException>(() => run.Text = "New");
            Assert.Throws<NotSupportedException>(() => run.FontSize = 24);
            Assert.That(run.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void ReadOnlyDisposedAndDetachedTextCannotBeMutated()
        {
            using var presentation = new XMLSlideShow();
            var run = NewRun(presentation);
            using var package = OPCPackage.Open(new MemoryStream(Save(presentation)), true);
            using var readOnly = new XMLSlideShow(package);
            var saved = FirstRun(readOnly);
            var before = saved.ParentParagraph.ParentShape.GetXmlObject().OuterXml;
            Assert.Throws<InvalidOperationException>(() => saved.Text = "Readonly");
            Assert.Throws<InvalidOperationException>(() => saved.FontSize = 24);
            Assert.Throws<InvalidOperationException>(() => saved.FontFamily = "Arial");
            Assert.Throws<InvalidOperationException>(() => saved.ParentParagraph.AddLineBreak());
            Assert.Throws<InvalidOperationException>(() => saved.ParentParagraph.TextAlign = TextAlignment.Left);
            Assert.Throws<InvalidOperationException>(() => saved.ParentParagraph.ParentShape.SetText("Readonly"));
            Assert.Throws<InvalidOperationException>(() => readOnly.GetSlides()[0].CreateTextBox());
            Assert.That(saved.ParentParagraph.ParentShape.GetXmlObject().OuterXml, Is.EqualTo(before));
            run.GetXmlObject().ParentNode.RemoveChild(run.GetXmlObject());
            Assert.Throws<InvalidOperationException>(() => run.IsBold = true);
            presentation.Dispose();
            Assert.Throws<ObjectDisposedException>(() => run.ParentParagraph.AddNewTextRun());
        }

        [TestCase("sample.pptx")]
        [TestCase("shapes.pptx")]
        public void EditingRealRegularRunsKeepsOtherPartsAndExistingFormattingBytes(string file)
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(file);
            using var presentation = new XMLSlideShow(input);
            var run = presentation.GetSlides().SelectMany(Shapes).SelectMany(shape => shape.TextParagraphs)
                .SelectMany(paragraph => paragraph.TextRuns).First(item => item.GetXmlObject().LocalName == "r");
            var shape = run.ParentParagraph.ParentShape;
            var properties = run.GetXmlObject()["rPr", A]?.OuterXml;
            var original = run.Text;
            var part = shape.GetSheet().GetPackagePart();
            var dependencies = presentation.Package.GetParts().Where(item => !item.IsRelationshipPart && item != part &&
                item.ContentType != XSLFRelation.MAIN.ContentType && !item.PartName.Name.StartsWith("/docProps/", StringComparison.Ordinal))
                .ToDictionary(item => item.PartName.Name, ReadBytes);
            run.Text = original + " \u4e2d\u6587";
            using var reopened = Open(Save(presentation));
            var saved = reopened.GetSlides().SelectMany(Shapes).Where(item => item.GetSheet().GetPackagePart().PartName.Equals(part.PartName))
                .Single(item => item.ID == shape.ID).TextParagraphs.SelectMany(paragraph => paragraph.TextRuns)
                .Single(item => item.GetXmlObject().LocalName == "r" && item.Text == original + " \u4e2d\u6587");
            Assert.That(saved.GetXmlObject()["rPr", A]?.OuterXml, Is.EqualTo(properties));
            foreach (var dependency in dependencies)
                Assert.That(ReadBytes(reopened.Package.GetPart(PackagingUriHelper.CreatePartName(dependency.Key))),
                    Is.EqualTo(dependency.Value), dependency.Key);
        }

        private static IXSLFShapeContainer Container(XMLSlideShow presentation, string kind)
        {
            if (kind == "master")
                return presentation.GetSlideMasters()[0];
            if (kind == "layout")
                return presentation.GetSlideMasters()[0].GetSlideLayouts()[0];
            var slide = presentation.GetSlides().FirstOrDefault() ?? presentation.CreateSlide();
            if (kind == "group")
                return slide.GetShapes().OfType<XSLFGroupShape>().FirstOrDefault() ?? slide.CreateGroup();
            return slide;
        }

        private static IEnumerable<XSLFSimpleShape> Shapes(IXSLFShapeContainer container)
        {
            foreach (var shape in container.GetShapes())
            {
                if (shape is XSLFSimpleShape simple)
                    yield return simple;
                if (shape is XSLFGroupShape group)
                    foreach (var child in Shapes(group))
                        yield return child;
            }
        }

        private static XSLFTextRun NewRun(XMLSlideShow presentation)
        {
            return presentation.CreateSlide().CreateTextBox().TextParagraphs[0].AddNewTextRun();
        }

        private static XSLFTextRun FirstRun(XMLSlideShow presentation)
        {
            return ((XSLFSimpleShape)presentation.GetSlides()[0].GetShapes()[0]).TextParagraphs[0].TextRuns[0];
        }

        private static void Placeholder(XSLFSimpleShape shape, string type, string index)
        {
            var parent = shape.GetXmlObject()["nvSpPr", P]["nvPr", P];
            parent.RemoveAll();
            Append(parent, "<p:ph type=\"" + type + "\" idx=\"" + index + "\"/>");
        }

        private static void ListStyle(XSLFSimpleShape shape, string content)
        {
            var style = shape.GetXmlObject()["txBody", P]["lstStyle", A];
            style.RemoveAll();
            if (content.Length != 0)
            {
                var root = Xml("<root>" + content + "</root>");
                foreach (XmlNode node in root.ChildNodes)
                    style.AppendChild(style.OwnerDocument.ImportNode(node, true));
            }
        }

        private static XmlElement Xml(string content)
        {
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml("<root xmlns:p=\"" + P + "\" xmlns:a=\"" + A + "\">" + content + "</root>");
            return (XmlElement)document.DocumentElement.FirstChild;
        }

        private static void Append(XmlElement parent, string content, bool prepend = false)
        {
            var element = parent.OwnerDocument.ImportNode(Xml(content), true);
            if (prepend)
                parent.PrependChild(element);
            else
                parent.AppendChild(element);
        }

        private static byte[] ReadBytes(PackagePart part)
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

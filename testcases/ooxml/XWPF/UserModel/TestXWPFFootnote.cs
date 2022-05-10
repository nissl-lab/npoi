using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.XWPF.UserModel
{
    [TestFixture]
    public class TestXWPFFootnote
    {
        private XWPFDocument docOut;
        private String p1Text;
        private String p2Text;
        private int footnoteId;
        private XWPFFootnote footnote;

        [SetUp]
        public void SetUp()
        {
            docOut = new XWPFDocument();
            p1Text = "First paragraph in footnote";
            p2Text = "Second paragraph in footnote";

            // NOTE: XWPFDocument.CreateFootnote() delegates directly
            //       to XWPFFootnotes.CreateFootnote() so this tests
            //       both creation of new XWPFFootnotes in document
            //       and XWPFFootnotes.CreateFootnote();

            // NOTE: Creating the footnote does not automatically
            //       create a first paragraph.
            footnote = docOut.CreateFootnote();
            footnoteId = footnote.Id;
        }
        [Test]
        public void TestAddParagraphsToFootnote()
        {

            // Add a run to the first paragraph:    

            XWPFParagraph p1 = footnote.CreateParagraph();
            p1.CreateRun().SetText(p1Text);

            // Create a second paragraph:

            XWPFParagraph p = footnote.CreateParagraph();
            Assert.IsNotNull(p, "Paragraph is null");
            p.CreateRun().SetText(p2Text);

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            XWPFFootnote testFootnote = docIn.GetFootnoteByID(footnoteId);
            Assert.IsNotNull(testFootnote);

            Assert.AreEqual(2, testFootnote.GetParagraphs().Count);
            XWPFParagraph testP1 = testFootnote.GetParagraphs()[0];
            Assert.AreEqual(p1Text, testP1.Text);

            XWPFParagraph testP2 = testFootnote.GetParagraphs()[1];
            Assert.AreEqual(p2Text, testP2.Text);

            // The first paragraph added using CreateParagraph() should
            // have the required footnote reference added to the first
            // run.

            // Verify that we have a footnote reference in the first paragraph and not
            // in the second paragraph.

            XWPFRun r1 = testP1.Runs[0];
            Assert.IsNotNull(r1);
            Assert.IsTrue(r1.GetCTR().GetFootnoteRefList().Count > 0, "No footnote reference in testP1");
            Assert.IsNotNull(r1.GetCTR().GetFootnoteRefArray(0), "No footnote reference in testP1");

            XWPFRun r2 = testP2.Runs[0];
            Assert.IsNotNull(r2, "Expected a run in testP2");
            Assert.IsTrue(r2.GetCTR().GetFootnoteRefList().Count == 0, "Found a footnote reference in testP2");

        }
        [Test]
        public void TestAddTableToFootnote()
        {
            XWPFTable table = footnote.CreateTable();
            Assert.IsNotNull(table);

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            XWPFFootnote testFootnote = docIn.GetFootnoteByID(footnoteId);
            XWPFTable testTable = testFootnote.GetTableArray(0);
            Assert.IsNotNull(testTable);

            table = footnote.CreateTable(2, 3);
            Assert.AreEqual(2, table.NumberOfRows);
            Assert.AreEqual(3, table.GetRow(0).GetTableCells().Count);

            // If the table is the first body element of the footnote then
            // a paragraph with the footnote reference should have been
            // added automatically.

            Assert.AreEqual(3, footnote.BodyElements.Count, "Expected 3 body elements");
            IBodyElement testP1 = footnote.BodyElements[0];
            Assert.IsTrue(testP1 is XWPFParagraph, "Expected a paragraph, got " + testP1.GetType().Name);
            XWPFRun r1 = ((XWPFParagraph)testP1).Runs[0];
            Assert.IsNotNull(r1);
            Assert.IsTrue(r1.GetCTR().GetFootnoteRefList().Count > 0, "No footnote reference in testP1");
            Assert.IsNotNull(r1.GetCTR().GetFootnoteRefArray(0), "No footnote reference in testP1");

        }
        [Test]
        public void TestRemoveFootnote()
        {
            // NOTE: XWPFDocument.removeFootnote() delegates directly to 
            //       XWPFFootnotes.
            docOut.CreateFootnote();
            Assert.AreEqual(2, docOut.GetFootnotes().Count, "Expected 2 footnotes");
            Assert.IsNotNull(docOut.GetFootnotes()[1], "Didn't get second footnote");
            bool result = docOut.RemoveFootnote(0);
            Assert.IsTrue(result, "Remove footnote did not return true");
            Assert.AreEqual(1, docOut.GetFootnotes().Count, "Expected 1 footnote after removal");
        }
        [Test]
        public void TestAddFootnoteRefToParagraph()
        {
            XWPFParagraph p = docOut.CreateParagraph();
            var runs = p.Runs;
            Assert.AreEqual(0, runs.Count, "Expected no runs in new paragraph");
            p.AddFootnoteReference(footnote);
            XWPFRun run = p.Runs[0];
            CT_R ctr = run.GetCTR();
            Assert.IsNotNull(run, "Expected a run");
            CT_FtnEdnRef ref1 = ctr.GetFootnoteReferenceList()[0];
            Assert.IsNotNull(ref1);
            Assert.AreEqual(footnote.Id.ToString(), ref1.id, "Footnote ID and reference ID did not match");
        }
    }
}
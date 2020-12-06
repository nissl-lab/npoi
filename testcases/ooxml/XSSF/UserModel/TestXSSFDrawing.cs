/* ===================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
using System.Collections.Generic;
using NUnit.Framework;
using System;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Dml;
using NPOI.Util;
using System.Drawing;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System.Text;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NPOI;

namespace TestCases.XSSF.UserModel
{
    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFDrawing
    {
        [Test]
        public void TestRead()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithDrawing.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            //the sheet has one relationship and it is XSSFDrawing
            List<POIXMLDocumentPart.RelationPart> rels = sheet.RelationParts;
            Assert.AreEqual(1, rels.Count);
            POIXMLDocumentPart.RelationPart rp = rels[0];
            Assert.IsTrue(rp.DocumentPart is XSSFDrawing);

            XSSFDrawing drawing = (XSSFDrawing)rp.DocumentPart;
            //sheet.CreateDrawingPatriarch() should return the same instance of XSSFDrawing
            Assert.AreSame(drawing, sheet.CreateDrawingPatriarch());
            String drawingId = rp.Relationship.Id;

            //there should be a relation to this Drawing in the worksheet
            Assert.IsTrue(sheet.GetCTWorksheet().IsSetDrawing());
            Assert.AreEqual(drawingId, sheet.GetCTWorksheet().drawing.id);


            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(6, shapes.Count);

            Assert.IsTrue(shapes[(0)] is XSSFPicture);
            Assert.IsTrue(shapes[(1)] is XSSFPicture);
            Assert.IsTrue(shapes[(2)] is XSSFPicture);
            Assert.IsTrue(shapes[(3)] is XSSFPicture);
            Assert.IsTrue(shapes[(4)] is XSSFSimpleShape);
            Assert.IsTrue(shapes[(5)] is XSSFPicture);

            foreach (XSSFShape sh in shapes)
                Assert.IsNotNull(sh.GetAnchor());

            checkRewrite(wb);
            wb.Close();
        }
        [Test]
        public void TestNew()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb1.CreateSheet();
            //multiple calls of CreateDrawingPatriarch should return the same instance of XSSFDrawing
            XSSFDrawing dr1 = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFDrawing dr2 = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            Assert.AreSame(dr1, dr2);

            List<POIXMLDocumentPart.RelationPart> rels = sheet.RelationParts;
            Assert.AreEqual(1, rels.Count);
            POIXMLDocumentPart.RelationPart rp = rels[0];
            Assert.IsTrue(rp.DocumentPart is XSSFDrawing);

            XSSFDrawing drawing = (XSSFDrawing)rp.DocumentPart;
            String drawingId = rp.Relationship.Id;

            //there should be a relation to this Drawing in the worksheet
            Assert.IsTrue(sheet.GetCTWorksheet().IsSetDrawing());
            Assert.AreEqual(drawingId, sheet.GetCTWorksheet().drawing.id);

            //XSSFClientAnchor anchor = new XSSFClientAnchor();

            XSSFConnector c1 = drawing.CreateConnector(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 2, 2));
            c1.LineWidth = 2.5;
            c1.LineStyle = LineStyle.DashDotSys;

            XSSFShapeGroup c2 = drawing.CreateGroup(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 5, 5));
            Assert.IsNotNull(c2);

            XSSFSimpleShape c3 = drawing.CreateSimpleShape(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));
            c3.SetText(new XSSFRichTextString("Test String"));
            c3.SetFillColor(128, 128, 128);

            XSSFTextBox c4 = (XSSFTextBox)drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 4, 4, 5, 6));
            XSSFRichTextString rt = new XSSFRichTextString("Test String");
            rt.ApplyFont(0, 5, wb1.CreateFont());
            rt.ApplyFont(5, 6, wb1.CreateFont());
            c4.SetText(rt);

            c4.IsNoFill = (true);

            Assert.AreEqual(4, drawing.GetCTDrawing().SizeOfTwoCellAnchorArray());

            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(4, shapes.Count);
            Assert.IsTrue(shapes[(0)] is XSSFConnector);
            Assert.IsTrue(shapes[(1)] is XSSFShapeGroup);
            Assert.IsTrue(shapes[(2)] is XSSFSimpleShape);
            Assert.IsTrue(shapes[(3)] is XSSFSimpleShape);

            // Save and re-load it
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();

            sheet = wb2.GetSheetAt(0) as XSSFSheet;

            // Check
            dr1 = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            CT_Drawing ctDrawing = dr1.GetCTDrawing();

            // Connector, shapes and text boxes are all two cell anchors
            Assert.AreEqual(0, ctDrawing.SizeOfAbsoluteAnchorArray());
            Assert.AreEqual(0, ctDrawing.SizeOfOneCellAnchorArray());
            Assert.AreEqual(4, ctDrawing.SizeOfTwoCellAnchorArray());

            shapes = dr1.GetShapes();
            Assert.AreEqual(4, shapes.Count);
            Assert.IsTrue(shapes[0] is XSSFConnector);
            Assert.IsTrue(shapes[1] is XSSFShapeGroup);
            Assert.IsTrue(shapes[2] is XSSFSimpleShape);
            Assert.IsTrue(shapes[3] is XSSFSimpleShape); //

            // Ensure it got the right namespaces
            //String xml = ctDrawing.ToString();
            //Assert.IsTrue(xml.Contains("xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\""));
            //Assert.IsTrue(xml.Contains("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""));

            checkRewrite(wb2);
            wb2.Close();
        }
        [Test]
        public void TestMultipleDrawings()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            for (int i = 0; i < 3; i++)
            {
                XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
                XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                Assert.IsNotNull(drawing);
            }
            OPCPackage pkg = wb.Package;
            try
            {
                Assert.AreEqual(3, pkg.GetPartsByContentType(XSSFRelation.DRAWINGS.ContentType).Count);
                checkRewrite(wb);
            }
            finally
            {
                pkg.Close();
            }
            wb.Close();
        }
        [Test]
        public void TestClone()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithDrawing.xlsx");
            XSSFSheet sheet1 = wb.GetSheetAt(0) as XSSFSheet;

            XSSFSheet sheet2 = wb.CloneSheet(0) as XSSFSheet;

            //the source sheet has one relationship and it is XSSFDrawing
            List<POIXMLDocumentPart> rels1 = sheet1.GetRelations();
            Assert.AreEqual(1, rels1.Count);
            Assert.IsTrue(rels1[(0)] is XSSFDrawing);

            List<POIXMLDocumentPart> rels2 = sheet2.GetRelations();
            Assert.AreEqual(1, rels2.Count);
            Assert.IsTrue(rels2[(0)] is XSSFDrawing);

            XSSFDrawing drawing1 = (XSSFDrawing)rels1[0];
            XSSFDrawing drawing2 = (XSSFDrawing)rels2[0];
            Assert.AreNotSame(drawing1, drawing2);  // Drawing2 is a clone of Drawing1

            List<XSSFShape> shapes1 = drawing1.GetShapes();
            List<XSSFShape> shapes2 = drawing2.GetShapes();
            Assert.AreEqual(shapes1.Count, shapes2.Count);

            for (int i = 0; i < shapes1.Count; i++)
            {
                XSSFShape sh1 = (XSSFShape)shapes1[(i)];
                XSSFShape sh2 = (XSSFShape)shapes2[i];

                Assert.IsTrue(sh1.GetType() == sh2.GetType());
                Assert.AreEqual(sh1.GetShapeProperties().ToString(), sh2.GetShapeProperties().ToString());
            }

            checkRewrite(wb);
            wb.Close();
        }

        /**
         * ensure that rich text attributes defined in a XSSFRichTextString
         * are passed to XSSFSimpleShape.
         *
         * See Bugzilla 52219.
         */
        [Test]
        public void TestRichText()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4)) as XSSFTextBox;
            XSSFRichTextString rt = new XSSFRichTextString("Test String");

            XSSFFont font = wb.CreateFont() as XSSFFont;
            font.SetColor(new XSSFColor(Color.FromArgb(0, 128, 128)));
            font.IsItalic = (true);
            font.IsBold = (true);
            font.Underline = FontUnderlineType.Single;
            rt.ApplyFont(font);

            shape.SetText(rt);

            CT_TextParagraph pr = shape.GetCTShape().txBody.p[0];
            Assert.AreEqual(1, pr.SizeOfRArray());

            CT_TextCharacterProperties rPr = pr.r[0].rPr;
            Assert.AreEqual(true, rPr.b);
            Assert.AreEqual(true, rPr.i);
            Assert.AreEqual(ST_TextUnderlineType.sng, rPr.u);
            Assert.IsTrue(Arrays.Equals(
                    new byte[] { 0, (byte)128, (byte)128 },
                    rPr.solidFill.srgbClr.val));

            checkRewrite(wb);
            wb.Close();
        }

        /**
         *  Test that anchor is not null when Reading shapes from existing Drawings
         */
        [Test]
        public void TestReadAnchors()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = wb1.CreateSheet() as XSSFSheet;
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4);
            XSSFShape shape1 = Drawing.CreateTextbox(anchor1) as XSSFShape;
            Assert.IsNotNull(shape1);

            XSSFClientAnchor anchor2 = new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 5);
            XSSFShape shape2 = Drawing.CreateTextbox(anchor2) as XSSFShape;
            Assert.IsNotNull(shape2);

            int pictureIndex = wb1.AddPicture(new byte[] { }, XSSFWorkbook.PICTURE_TYPE_PNG);
            XSSFClientAnchor anchor3 = new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 6);
            XSSFShape shape3 = Drawing.CreatePicture(anchor3, pictureIndex) as XSSFShape;
            Assert.IsNotNull(shape3);

            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as XSSFSheet;
            Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            List<XSSFShape> shapes = Drawing.GetShapes();
            Assert.AreEqual(3, shapes.Count);
            Assert.AreEqual(shapes[0].GetAnchor(), anchor1);
            Assert.AreEqual(shapes[1].GetAnchor(), anchor2);
            Assert.AreEqual(shapes[2].GetAnchor(), anchor3);

            checkRewrite(wb2);
            wb2.Close();
        }

        /**
         * ensure that font and color rich text attributes defined in a XSSFRichTextString
         * are passed to XSSFSimpleShape.
         *
         * See Bugzilla 54969.
         */

        [Test]
        public void TestRichTextFontAndColor()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4)) as XSSFTextBox;
            XSSFRichTextString rt = new XSSFRichTextString("Test String");

            XSSFFont font = wb.CreateFont() as XSSFFont;
            font.SetColor(new XSSFColor(Color.FromArgb(0, 128, 128)));
            font.FontName = ("Arial");
            rt.ApplyFont(font);

            shape.SetText(rt);

            CT_TextParagraph pr = shape.GetCTShape().txBody.GetPArray(0);
            Assert.AreEqual(1, pr.SizeOfRArray());

            CT_TextCharacterProperties rPr = pr.GetRArray(0).rPr;
            Assert.AreEqual("Arial", rPr.latin.typeface);
            Assert.IsTrue(Arrays.Equals(
                    new byte[] { 0, (byte)128, (byte)128 },
                    rPr.solidFill.srgbClr.val));

            checkRewrite(wb);
            wb.Close();
        }

        ///**
        // * Test SetText single paragraph to ensure backwards compatibility
        // */
        [Test]
        public void TestSetTextSingleParagraph()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFTextBox shape = drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));
            XSSFRichTextString rt = new XSSFRichTextString("Test String");

            XSSFFont font = wb.CreateFont() as XSSFFont;
            font.SetColor(new XSSFColor(Color.FromArgb(0, 255, 255)));
            font.FontName = ("Arial");
            rt.ApplyFont(font);

            shape.SetText(rt);

            List<XSSFTextParagraph> paras = shape.TextParagraphs;
            Assert.AreEqual(1, paras.Count);
            Assert.AreEqual("Test String", paras[0].Text);

            List<XSSFTextRun> runs = paras[0].TextRuns;
            Assert.AreEqual(1, runs.Count);
            Assert.AreEqual("Arial", runs[0].FontFamily);

            Color clr = runs[0].FontColor;
            Assert.IsTrue(Arrays.Equals(
                    new int[] { 0, 255, 255 },
                    new int[] { clr.R, clr.G, clr.B }));

            checkRewrite(wb);
            wb.Close();
        }

        ///**
        // * Test AddNewTextParagraph 
        // */
        [Test]
        public void TestAddNewTextParagraph()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFTextBox shape = drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));

            XSSFTextParagraph para = shape.AddNewTextParagraph();
            para.AddNewTextRun().Text = ("Line 1");

            List<XSSFTextParagraph> paras = shape.TextParagraphs;
            Assert.AreEqual(2, paras.Count);	// this should be 2 as XSSFSimpleShape Creates a default paragraph (no text), and then we add a string to that.

            List<XSSFTextRun> runs = para.TextRuns;
            Assert.AreEqual(1, runs.Count);
            Assert.AreEqual("Line 1", runs[0].Text);

            checkRewrite(wb);
            wb.Close();
        }

        ///**
        // * Test AddNewTextParagraph using RichTextString
        // */
        [Test]
        public void TestAddNewTextParagraphWithRTS()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = wb1.CreateSheet() as XSSFSheet;
            XSSFDrawing drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFTextBox shape = drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));
            XSSFRichTextString rt = new XSSFRichTextString("Test Rich Text String");

            XSSFFont font = wb1.CreateFont() as XSSFFont;
            font.SetColor(new XSSFColor(Color.FromArgb(0, 255, 255)));
            font.FontName = ("Arial");
            rt.ApplyFont(font);

            XSSFFont midfont = wb1.CreateFont() as XSSFFont;
            midfont.SetColor(new XSSFColor(Color.FromArgb(0, 255, 0)));
            rt.ApplyFont(5, 14, midfont);	// Set the text "Rich Text" to be green and the default font

            XSSFTextParagraph para = shape.AddNewTextParagraph(rt);

            // Save and re-load it
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as XSSFSheet;

            // Check
            drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(1, shapes.Count);
            Assert.IsTrue(shapes[0] is XSSFSimpleShape);

            XSSFSimpleShape sshape = (XSSFSimpleShape)shapes[0];

            List<XSSFTextParagraph> paras = sshape.TextParagraphs;
            Assert.AreEqual(2, paras.Count);	// this should be 2 as XSSFSimpleShape Creates a default paragraph (no text), and then we add a string to that.  

            List<XSSFTextRun> runs = para.TextRuns;
            Assert.AreEqual(3, runs.Count);

            // first run properties
            Assert.AreEqual("Test ", runs[0].Text);
            Assert.AreEqual("Arial", runs[0].FontFamily);

            Color clr = runs[0].FontColor;
            Assert.IsTrue(Arrays.Equals(
                    new int[] { 0, 255, 255 },
                    new int[] { clr.R, clr.G, clr.B }));

            // second run properties        
            Assert.AreEqual("Rich Text", runs[1].Text);
            Assert.AreEqual(XSSFFont.DEFAULT_FONT_NAME, runs[1].FontFamily);

            clr = runs[1].FontColor;
            Assert.IsTrue(Arrays.Equals(
                    new int[] { 0, 255, 0 },
                    new int[] { clr.R, clr.G, clr.B }));

            // third run properties
            Assert.AreEqual(" String", runs[2].Text);
            Assert.AreEqual("Arial", runs[2].FontFamily);
            clr = runs[2].FontColor;
            Assert.IsTrue(Arrays.Equals(
                    new int[] { 0, 255, 255 },
                    new int[] { clr.R, clr.G, clr.B }));

            checkRewrite(wb2);
            wb2.Close();
        }

        ///**
        // * Test add multiple paragraphs and retrieve text
        // */
        [Test]
        public void TestAddMultipleParagraphs()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing; ;

            XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));

            XSSFTextParagraph para = shape.AddNewTextParagraph();
            para.AddNewTextRun().Text = ("Line 1");

            para = shape.AddNewTextParagraph();
            para.AddNewTextRun().Text = ("Line 2");

            para = shape.AddNewTextParagraph();
            para.AddNewTextRun().Text = ("Line 3");

            List<XSSFTextParagraph> paras = shape.TextParagraphs;
            Assert.AreEqual(4, paras.Count);	// this should be 4 as XSSFSimpleShape Creates a default paragraph (no text), and then we Added 3 paragraphs
            Assert.AreEqual("Line 1\nLine 2\nLine 3", shape.Text);

            checkRewrite(wb);
            wb.Close();
        }

        ///**
        // * Test Setting the text, then Adding multiple paragraphs and retrieve text
        // */
        [Test]
        public void TestSetAddMultipleParagraphs()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing; ;

            XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));

            shape.SetText("Line 1");

            XSSFTextParagraph para = shape.AddNewTextParagraph();
            para.AddNewTextRun().Text = ("Line 2");

            para = shape.AddNewTextParagraph();
            para.AddNewTextRun().Text = ("Line 3");

            List<XSSFTextParagraph> paras = shape.TextParagraphs;
            Assert.AreEqual(3, paras.Count);	// this should be 3 as we overwrote the default paragraph with SetText, then Added 2 new paragraphs
            Assert.AreEqual("Line 1\nLine 2\nLine 3", shape.Text);

            checkRewrite(wb);
            wb.Close();
        }

        ///**
        // * Test Reading text from a textbox in an existing file
        // */
        [Test]
        public void TestReadTextBox()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithDrawing.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            //the sheet has one relationship and it is XSSFDrawing
            List<POIXMLDocumentPart.RelationPart> rels = sheet.RelationParts;
            Assert.AreEqual(1, rels.Count);
            POIXMLDocumentPart.RelationPart rp = rels[0];
            Assert.IsTrue(rp.DocumentPart is XSSFDrawing);

            XSSFDrawing drawing = (XSSFDrawing)rp.DocumentPart;
            //sheet.CreateDrawingPatriarch() should return the same instance of XSSFDrawing
            Assert.AreSame(drawing, sheet.CreateDrawingPatriarch());
            String drawingId = rp.Relationship.Id;

            //there should be a relation to this Drawing in the worksheet
            Assert.IsTrue(sheet.GetCTWorksheet().IsSetDrawing());
            Assert.AreEqual(drawingId, sheet.GetCTWorksheet().drawing.id);

            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(6, shapes.Count);

            Assert.IsTrue(shapes[4] is XSSFSimpleShape);

            XSSFSimpleShape textbox = (XSSFSimpleShape)shapes[4];
            Assert.AreEqual("Sheet with various pictures\n(jpeg, png, wmf, emf and pict)", textbox.Text);

            checkRewrite(wb);
            wb.Close();
        }


        ///**
        // * Test Reading multiple paragraphs from a textbox in an existing file
        // */
        [Test]
        public void TestReadTextBoxParagraphs()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithTextBox.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            //the sheet has one relationship and it is XSSFDrawing
            List<POIXMLDocumentPart.RelationPart> rels = sheet.RelationParts;
            Assert.AreEqual(1, rels.Count);
            POIXMLDocumentPart.RelationPart rp = rels[0];
            Assert.IsTrue(rp.DocumentPart is XSSFDrawing);

            XSSFDrawing drawing = (XSSFDrawing)rp.DocumentPart;

            //sheet.CreateDrawingPatriarch() should return the same instance of XSSFDrawing
            Assert.AreSame(drawing, sheet.CreateDrawingPatriarch());
            String drawingId = rp.Relationship.Id;

            //there should be a relation to this Drawing in the worksheet
            Assert.IsTrue(sheet.GetCTWorksheet().IsSetDrawing());
            Assert.AreEqual(drawingId, sheet.GetCTWorksheet().drawing.id);

            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(1, shapes.Count);

            Assert.IsTrue(shapes[0] is XSSFSimpleShape);

            XSSFSimpleShape textbox = (XSSFSimpleShape)shapes[0];

            List<XSSFTextParagraph> paras = textbox.TextParagraphs;
            Assert.AreEqual(3, paras.Count);

            Assert.AreEqual("Line 2", paras[1].Text);	// check content of second paragraph

            Assert.AreEqual("Line 1\nLine 2\nLine 3", textbox.Text);	// check content of entire textbox

            // check attributes of paragraphs
            Assert.AreEqual(TextAlign.LEFT, paras[0].TextAlign);
            Assert.AreEqual(TextAlign.CENTER, paras[1].TextAlign);
            Assert.AreEqual(TextAlign.RIGHT, paras[2].TextAlign);

            Color clr = paras[0].TextRuns[0].FontColor;
            Assert.IsTrue(Arrays.Equals(
                    new int[] { 255, 0, 0 },
                    new int[] { clr.R, clr.G, clr.B }));

            clr = paras[1].TextRuns[0].FontColor;
            Assert.IsTrue(Arrays.Equals(
                    new int[] { 0, 255, 0 },
                    new int[] { clr.R, clr.G, clr.B }));

            clr = paras[2].TextRuns[0].FontColor;
            Assert.IsTrue(Arrays.Equals(
                    new int[] { 0, 0, 255 },
                    new int[] { clr.R, clr.G, clr.B }));

            checkRewrite(wb);
            wb.Close();
        }

        /**
         * Test Adding and Reading back paragraphs as bullet points
         */
        [Test]
        public void TestAddBulletParagraphs()
        {

            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = wb1.CreateSheet() as XSSFSheet;
            XSSFDrawing drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFTextBox shape = drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 10, 20));

            String paraString1 = "A normal paragraph";
            String paraString2 = "First bullet";
            String paraString3 = "Second bullet (level 1)";
            String paraString4 = "Third bullet";
            String paraString5 = "Another normal paragraph";
            String paraString6 = "First numbered bullet";
            String paraString7 = "Second bullet (level 1)";
            String paraString8 = "Third bullet (level 1)";
            String paraString9 = "Fourth bullet (level 1)";
            String paraString10 = "Fifth Bullet";

            XSSFTextParagraph para = shape.AddNewTextParagraph(paraString1);
            para = shape.AddNewTextParagraph(paraString2);
            para.SetBullet(true);

            para = shape.AddNewTextParagraph(paraString3);
            para.SetBullet(true);
            para.Level = (1);

            para = shape.AddNewTextParagraph(paraString4);
            para.SetBullet(true);

            para = shape.AddNewTextParagraph(paraString5);
            para = shape.AddNewTextParagraph(paraString6);
            para.SetBullet(ListAutoNumber.ARABIC_PERIOD);

            para = shape.AddNewTextParagraph(paraString7);
            para.SetBullet(ListAutoNumber.ARABIC_PERIOD, 3);
            para.Level = (1);

            para = shape.AddNewTextParagraph(paraString8);
            para.SetBullet(ListAutoNumber.ARABIC_PERIOD, 3);
            para.Level = (1);

            para = shape.AddNewTextParagraph("");
            para.SetBullet(ListAutoNumber.ARABIC_PERIOD, 3);
            para.Level = (1);

            para = shape.AddNewTextParagraph(paraString9);
            para.SetBullet(ListAutoNumber.ARABIC_PERIOD, 3);
            para.Level = (1);

            para = shape.AddNewTextParagraph(paraString10);
            para.SetBullet(ListAutoNumber.ARABIC_PERIOD);

            // Save and re-load it
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            sheet = wb2.GetSheetAt(0) as XSSFSheet;

            // Check
            drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(1, shapes.Count);
            Assert.IsTrue(shapes[0] is XSSFSimpleShape);

            XSSFSimpleShape sshape = (XSSFSimpleShape)shapes[0];

            List<XSSFTextParagraph> paras = sshape.TextParagraphs;
            Assert.AreEqual(12, paras.Count);  // this should be 12 as XSSFSimpleShape Creates a default paragraph (no text), and then we Added to that

            StringBuilder builder = new StringBuilder();

            builder.Append(paraString1);
            builder.Append("\n");
            builder.Append("\u2022 ");
            builder.Append(paraString2);
            builder.Append("\n");
            builder.Append("\t\u2022 ");
            builder.Append(paraString3);
            builder.Append("\n");
            builder.Append("\u2022 ");
            builder.Append(paraString4);
            builder.Append("\n");
            builder.Append(paraString5);
            builder.Append("\n");
            builder.Append("1. ");
            builder.Append(paraString6);
            builder.Append("\n");
            builder.Append("\t3. ");
            builder.Append(paraString7);
            builder.Append("\n");
            builder.Append("\t4. ");
            builder.Append(paraString8);
            builder.Append("\n");
            builder.Append("\t");   // should be empty
            builder.Append("\n");
            builder.Append("\t5. ");
            builder.Append(paraString9);
            builder.Append("\n");
            builder.Append("2. ");
            builder.Append(paraString10);

            Assert.AreEqual(builder.ToString(), sshape.Text);

            checkRewrite(wb2);
            wb2.Close();
        }

        ///**
        // * Test Reading bullet numbering from a textbox in an existing file
        // */
        [Test]
        public void TestReadTextBox2()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithTextBox2.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFDrawing drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            List<XSSFShape> shapes = drawing.GetShapes();
            XSSFSimpleShape textbox = (XSSFSimpleShape)shapes[0];
            String extracted = textbox.Text;
            StringBuilder sb = new StringBuilder();
            sb.Append("1. content1A\n");
            sb.Append("\t1. content1B\n");
            sb.Append("\t2. content2B\n");
            sb.Append("\t3. content3B\n");
            sb.Append("2. content2A\n");
            sb.Append("\t3. content2BStartAt3\n");
            sb.Append("\t\n\t\n\t");
            sb.Append("4. content2BStartAt3Incremented\n");
            sb.Append("\t\n\t\n\t\n\t");

            Assert.AreEqual(sb.ToString(), extracted);

            checkRewrite(wb);
            wb.Close();
        }

        [Test]
        public void TestXSSFSimpleShapeCausesNPE56514()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("56514.xlsx");
            XSSFSheet sheet = wb1.GetSheetAt(0) as XSSFSheet;
            XSSFDrawing drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(4, shapes.Count);

            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as XSSFSheet;
            drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            shapes = drawing.GetShapes();
            Assert.AreEqual(4, shapes.Count);
            wb2.Close();
        }


        [Test]
        public void TestBug56835CellComment()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
                XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

                // first comment works
                IClientAnchor anchor = new XSSFClientAnchor(1, 1, 2, 2, 3, 3, 4, 4);
                XSSFComment comment = Drawing.CreateCellComment(anchor) as XSSFComment;
                Assert.IsNotNull(comment);

                try
                {
                    Drawing.CreateCellComment(anchor);
                    Assert.Fail("Should fail if we try to add the same comment for the same cell");
                }
                catch (ArgumentException e)
                {
                    // expected
                }
            }
            finally
            {
                wb.Close();
            }
        }
        private static void checkRewrite(XSSFWorkbook wb)
        {
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wb2);
            wb2.Close();
        }
    }
}


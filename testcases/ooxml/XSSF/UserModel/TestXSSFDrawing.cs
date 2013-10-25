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
namespace NPOI.XSSF.UserModel
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
            List<POIXMLDocumentPart> rels = sheet.GetRelations();
            Assert.AreEqual(1, rels.Count);
            Assert.IsTrue(rels[0] is XSSFDrawing);

            XSSFDrawing drawing = (XSSFDrawing)rels[0];
            //sheet.CreateDrawingPatriarch() should return the same instance of XSSFDrawing
            Assert.AreSame(drawing, sheet.CreateDrawingPatriarch());
            String drawingId = drawing.GetPackageRelationship().Id;

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

        }
        [Test]
        public void TestNew()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            //multiple calls of CreateDrawingPatriarch should return the same instance of XSSFDrawing
            XSSFDrawing dr1 = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFDrawing dr2 = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            Assert.AreSame(dr1, dr2);

            List<POIXMLDocumentPart> rels = sheet.GetRelations();
            Assert.AreEqual(1, rels.Count);
            Assert.IsTrue(rels[0] is XSSFDrawing);

            XSSFDrawing drawing = (XSSFDrawing)rels[0];
            String drawingId = drawing.GetPackageRelationship().Id;

            //there should be a relation to this Drawing in the worksheet
            Assert.IsTrue(sheet.GetCTWorksheet().IsSetDrawing());
            Assert.AreEqual(drawingId, sheet.GetCTWorksheet().drawing.id);

            //XSSFClientAnchor anchor = new XSSFClientAnchor();

            XSSFConnector c1 = drawing.CreateConnector(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 2, 2));
            c1.LineWidth = 2.5;
            c1.LineStyle = SS.UserModel.LineStyle.DashDotSys;

            XSSFShapeGroup c2 = drawing.CreateGroup(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 5, 5));

            XSSFSimpleShape c3 = drawing.CreateSimpleShape(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));
            c3.SetText(new XSSFRichTextString("Test String"));
            c3.SetFillColor(128, 128, 128);

            XSSFTextBox c4 = (XSSFTextBox)drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 4, 4, 5, 6));
            XSSFRichTextString rt = new XSSFRichTextString("Test String");
            rt.ApplyFont(0, 5, wb.CreateFont());
            rt.ApplyFont(5, 6, wb.CreateFont());
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
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sheet = wb.GetSheetAt(0) as XSSFSheet;

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
        }
        [Test]
        public void TestMultipleDrawings()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            for (int i = 0; i < 3; i++)
            {
                XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
                XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            }
            OPCPackage pkg = wb.Package;
            Assert.AreEqual(3, pkg.GetPartsByContentType(XSSFRelation.DRAWINGS.ContentType).Count);
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
            font.IsBold=(true);
            font.Underline=FontUnderlineType.Single;
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

        }

        /**
         *  Test that anchor is not null when Reading shapes from existing Drawings
         */
        [Test]
        public void TestReadAnchors()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

            XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4);
            XSSFShape shape1 = Drawing.CreateTextbox(anchor1) as XSSFShape;

            XSSFClientAnchor anchor2 = new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 5);
            XSSFShape shape2 = Drawing.CreateTextbox(anchor2) as XSSFShape;

            int pictureIndex = wb.AddPicture(new byte[] { }, XSSFWorkbook.PICTURE_TYPE_PNG);
            XSSFClientAnchor anchor3 = new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 6);
            XSSFShape shape3 = Drawing.CreatePicture(anchor3, pictureIndex) as XSSFShape;

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sheet = wb.GetSheetAt(0) as XSSFSheet;
            Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            List<XSSFShape> shapes = Drawing.GetShapes();
            Assert.AreEqual(3, shapes.Count);
            Assert.AreEqual(shapes[0].GetAnchor(), anchor1);
            Assert.AreEqual(shapes[1].GetAnchor(), anchor2);
            Assert.AreEqual(shapes[2].GetAnchor(), anchor3);


        }

    }
}


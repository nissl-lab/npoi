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

            XSSFClientAnchor anchor = new XSSFClientAnchor();

            XSSFConnector c1 = drawing.CreateConnector(anchor);
            c1.LineWidth=3;
            c1.LineStyle= SS.UserModel.LineStyle.DashDotSys;

            XSSFShapeGroup c2 = drawing.CreateGroup(anchor);

            XSSFSimpleShape c3 = drawing.CreateSimpleShape(anchor);
            c3.SetText(new XSSFRichTextString("Test String"));
            c3.SetFillColor(128, 128, 128);

            XSSFTextBox c4 = (XSSFTextBox)drawing.CreateTextbox(anchor);
            XSSFRichTextString rt = new XSSFRichTextString("Test String");
            rt.ApplyFont(0, 5, wb.CreateFont());
            rt.ApplyFont(5, 6, wb.CreateFont());
            c4.SetText(rt);

            c4.IsNoFill=(true);
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
    }
}


/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;

/*
 This sample is copied from poi.hssf.usermodel.examples. Original name is OfficeDrawing.java
 */

namespace DrawingInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.CreateSheet("new sheet");
            ISheet sheet2 = hssfworkbook.CreateSheet("second sheet");
            ISheet sheet3 = hssfworkbook.CreateSheet("third sheet");
            ISheet sheet4 = hssfworkbook.CreateSheet("fourth sheet");

            // Draw stuff in them
            DrawSheet1(sheet1);
            DrawSheet2(sheet2);
            DrawSheet3(sheet3);
            DrawSheet4(sheet4, hssfworkbook);

            WriteToFile();
        }


        private static void DrawSheet1(ISheet sheet1)
        {
            // Create a row and size one of the cells reasonably large.
            IRow row = sheet1.CreateRow(2);
            row.Height = ((short)2800);
            row.CreateCell(1);
            sheet1.SetColumnWidth(2, 9000);

            // Create the Drawing patriarch.  This is the top level container for
            // all shapes.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet1.CreateDrawingPatriarch();

            // Draw some lines and an oval.
            DrawLinesToCenter(patriarch);
            DrawManyLines(patriarch);
            DrawOval(patriarch);
            DrawPolygon(patriarch);

            // Draw a rectangle.
            HSSFSimpleShape rect = patriarch.CreateSimpleShape(new HSSFClientAnchor(100, 100, 900, 200, (short)0, 0, (short)0, 0));
            rect.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        }

        private static void DrawSheet2(ISheet sheet2)
        {
            // Create a row and size one of the cells reasonably large.
            IRow row = sheet2.CreateRow(2);
            row.CreateCell(1);
            row.HeightInPoints = 240;
            sheet2.SetColumnWidth(2, 9000);

            // Create the Drawing patriarch.  This is the top level container for
            // all shapes. This will clear out any existing shapes for that sheet.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet2.CreateDrawingPatriarch();

            // Draw a grid in one of the cells.
            DrawGrid(patriarch);
        }

        private static void DrawSheet3(ISheet sheet3)
        {
            // Create a row and size one of the cells reasonably large
            IRow row = sheet3.CreateRow(2);
            row.HeightInPoints = 140;
            row.CreateCell(1);
            sheet3.SetColumnWidth(2, 9000);

            // Create the Drawing patriarch.  This is the top level container for
            // all shapes. This will clear out any existing shapes for that sheet.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet3.CreateDrawingPatriarch();

            // Create a shape group.
            HSSFShapeGroup group = patriarch.CreateGroup(
                    new HSSFClientAnchor(0, 0, 900, 200, (short)2, 2, (short)2, 2));

            // Create a couple of lines in the group.
            HSSFSimpleShape shape1 = group.CreateShape(new HSSFChildAnchor(3, 3, 500, 500));
            shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            ((HSSFChildAnchor)shape1.Anchor).SetAnchor((short)3, 3, 500, 500);
            HSSFSimpleShape shape2 = group.CreateShape(new HSSFChildAnchor((short)1, 200, 400, 600));
            shape2.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);

        }

        private static void DrawSheet4(ISheet sheet4, HSSFWorkbook wb)
        {
            // Create the Drawing patriarch.  This is the top level container for
            // all shapes. This will clear out any existing shapes for that sheet.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet4.CreateDrawingPatriarch();

            // Create a couple of textboxes
            HSSFTextbox textbox1 = (HSSFTextbox)patriarch.CreateTextbox(
                    new HSSFClientAnchor(0, 0, 0, 0, (short)1, 1, (short)2, 2));
            textbox1.String = new HSSFRichTextString("This is a test");
            HSSFTextbox textbox2 = (HSSFTextbox)patriarch.CreateTextbox(
                    new HSSFClientAnchor(0, 0, 900, 100, (short)3, 3, (short)3, 4));
            textbox2.String = new HSSFRichTextString("Woo");
            textbox2.SetFillColor(200, 0, 0);
            textbox2.LineStyle = LineStyle.DotGel;
             
            // Create third one with some fancy font styling.
            HSSFTextbox textbox3 = (HSSFTextbox)patriarch.CreateTextbox(
                    new HSSFClientAnchor(0, 0, 900, 100, (short)4, 4, (short)5, 4 + 1));
            IFont font = wb.CreateFont();
            font.IsItalic = true;
            font.Underline = (byte)FontUnderlineType.DOUBLE;
            HSSFRichTextString str = new HSSFRichTextString("Woo!!!");
            str.ApplyFont(2, 5, font);
            textbox3.String = str;
            textbox3.FillColor = 0x08000030;
            textbox3.LineStyle = LineStyle.None;  // no line around the textbox.
            textbox3.IsNoFill = true;    // make it transparent
        }

        private static void DrawOval(HSSFPatriarch patriarch)
        {
            // Create an oval and style to taste.
            HSSFClientAnchor a = new HSSFClientAnchor();
            a.SetAnchor((short)2, 2, 20, 20, (short)2, 2, 190, 80);
            HSSFSimpleShape s = patriarch.CreateSimpleShape(a);
            s.ShapeType = HSSFSimpleShape.OBJECT_TYPE_OVAL;
            s.SetLineStyleColor(10, 10, 10);
            s.SetFillColor(90, 10, 200);
            s.LineWidth = HSSFShape.LINEWIDTH_ONE_PT * 3;
            s.LineStyle = LineStyle.DotSys;
        }

        private static void DrawPolygon(HSSFPatriarch patriarch)
        {
            HSSFClientAnchor a = new HSSFClientAnchor();
            a.SetAnchor((short)2, 2, 0, 0, (short)3, 3, 1023, 255);
            HSSFShapeGroup g = patriarch.CreateGroup(a);
            g.SetCoordinates(0, 0, 200, 200);
            HSSFPolygon p1 = g.CreatePolygon(new HSSFChildAnchor(0, 0, 200, 200));
            p1.SetPolygonDrawArea(100, 100);
            p1.SetPoints(new int[] { 0, 90, 50 }, new int[] { 5, 5, 44 });
            p1.SetFillColor(0, 255, 0);
            HSSFPolygon p2 = g.CreatePolygon(new HSSFChildAnchor(20, 20, 200, 200));
            p2.SetPolygonDrawArea(200, 200);
            p2.SetPoints(new int[] { 120, 20, 150 }, new int[] { 105, 30, 195 });
            p2.SetFillColor(255, 0, 0);
        }

        private static void DrawManyLines(HSSFPatriarch patriarch)
        {
            // Draw bunch of lines
            int x1 = 100;
            int y1 = 100;
            int x2 = 800;
            int y2 = 200;
            int color = 0;
            for (int i = 0; i < 10; i++)
            {
                HSSFClientAnchor a2 = new HSSFClientAnchor();
                a2.SetAnchor((short)2, 2, x1, y1, (short)2, 2, x2, y2);
                HSSFSimpleShape shape2 = patriarch.CreateSimpleShape(a2);
                shape2.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
                shape2.SetLineStyleColor(color);
                y1 -= 10;
                y2 -= 10;
                color += 30;
            }
        }

        private static void DrawGrid(HSSFPatriarch patriarch)
        {
            // This Draws a grid of lines.  Since the coordinates space fixed at
            // 1024 by 256 we use a ratio to get a reasonably square grids.

            double xRatio = 3.22;
            double yRatio = 0.6711;

            int x1 = 000;
            int y1 = 000;
            int x2 = 000;
            int y2 = 200;
            for (int i = 0; i < 20; i++)
            {
                HSSFClientAnchor a2 = new HSSFClientAnchor();
                a2.SetAnchor((short)2, 2, (int)(x1 * xRatio), (int)(y1 * yRatio),
                        (short)2, 2, (int)(x2 * xRatio), (int)(y2 * yRatio));
                HSSFSimpleShape shape2 = patriarch.CreateSimpleShape(a2);
                shape2.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);

                x1 += 10;
                x2 += 10;
            }

            x1 = 000;
            y1 = 000;
            x2 = 200;
            y2 = 000;
            for (int i = 0; i < 20; i++)
            {
                HSSFClientAnchor a2 = new HSSFClientAnchor();
                a2.SetAnchor((short)2, 2, (int)(x1 * xRatio), (int)(y1 * yRatio),
                        (short)2, 2, (int)(x2 * xRatio), (int)(y2 * yRatio));
                HSSFSimpleShape shape2 = patriarch.CreateSimpleShape(a2);
                shape2.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);

                y1 += 10;
                y2 += 10;
            }
        }

        private static void DrawLinesToCenter(HSSFPatriarch patriarch)
        {
            // Draw some lines from and to the corners
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)2, 2, 0, 0, (short)2, 2, 512, 128);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)2, 2, 512, 128, (short)2, 2, 1023, 0);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)1, 1, 0, 0, (short)1, 1, 512, 100);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)1, 1, 512, 100, (short)1, 1, 1023, 0);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }

        }

        static HSSFWorkbook hssfworkbook;

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();

            //Create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //Create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
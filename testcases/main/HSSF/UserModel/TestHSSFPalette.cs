/* ====================================================================
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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using System.Collections;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.HSSF.Util;
    using TestCases.HSSF;
    using NUnit.Framework;
    using NPOI.SS.UserModel;
    using SixLabors.ImageSharp.PixelFormats;

    /**
     * @author Brian Sanders (bsanders at risklabs dot com)
     */
    [TestFixture]
    public class TestHSSFPalette
    {
        private PaletteRecord palette;
        private HSSFPalette hssfPalette;



        [SetUp]
        public void SetUp()
        {
            palette = new PaletteRecord();
            hssfPalette = new HSSFPalette(palette);
        }

        /**
         * Verifies that a custom palette can be Created, saved, and reloaded
         */
        [Test]
        public void TestCustomPalette()
        {
            //reading sample xls
            HSSFWorkbook book = HSSFTestDataSamples.OpenSampleWorkbook("Simple.xls");

            //creating custom palette
            HSSFPalette palette = book.GetCustomPalette();
            palette.SetColorAtIndex((short)0x12, (byte)101, (byte)230, (byte)100);
            palette.SetColorAtIndex((short)0x3b, (byte)0, (byte)255, (byte)52);

            //writing to disk; reading in and verifying palette
            string tmppath = TempFile.GetTempFilePath("TestCustomPalette", ".xls");
            FileStream fos = new FileStream(tmppath, FileMode.OpenOrCreate);
            book.Write(fos);
            fos.Close();

            FileStream fis = new FileStream(tmppath, FileMode.Open,FileAccess.Read);
            book = new HSSFWorkbook(fis);
            fis.Close();

            palette = book.GetCustomPalette();
            HSSFColor color = palette.GetColor(HSSFColor.Coral.Index);  //unmodified
            Assert.IsNotNull(color, "Unexpected null in custom palette (unmodified index)");
            byte[] expectedRGB = HSSFColor.Coral.Triplet;
            byte[] actualRGB = color.RGB;
            String msg = "Expected palette position to remain unmodified";
            Assert.AreEqual(expectedRGB[0], actualRGB[0], msg);
            Assert.AreEqual(expectedRGB[1], actualRGB[1], msg);
            Assert.AreEqual(expectedRGB[2], actualRGB[2], msg);

            color = palette.GetColor((short)0x12);
            Assert.IsNotNull(color, "Unexpected null in custom palette (modified)");
            actualRGB = color.RGB;
            msg = "Expected palette modification to be preserved across save";
            Assert.AreEqual((short)101, actualRGB[0], msg);
            Assert.AreEqual((short)230, actualRGB[1], msg);
            Assert.AreEqual((short)100, actualRGB[2], msg);
        }

        /**
         * Uses the palette from cell stylings
         */
        [Test]
        public void TestPaletteFromCellColours()
        {
            HSSFWorkbook book = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithColours.xls");

            HSSFPalette p = book.GetCustomPalette();

            ICell cellA = book.GetSheetAt(0).GetRow(0).GetCell(0);
            ICell cellB = book.GetSheetAt(0).GetRow(1).GetCell(0);
            ICell cellC = book.GetSheetAt(0).GetRow(2).GetCell(0);
            ICell cellD = book.GetSheetAt(0).GetRow(3).GetCell(0);
            ICell cellE = book.GetSheetAt(0).GetRow(4).GetCell(0);

            // Plain
            Assert.AreEqual("I'm plain", cellA.StringCellValue);
            Assert.AreEqual(64, cellA.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, cellA.CellStyle.FillBackgroundColor);
            Assert.AreEqual(HSSFColor.COLOR_NORMAL, cellA.CellStyle.GetFont(book).Color);
            Assert.AreEqual(0, (short)cellA.CellStyle.FillPattern);
            Assert.AreEqual("0:0:0", p.GetColor((short)64).GetHexString());
            Assert.AreEqual(null, p.GetColor((short)32767));

            // Red
            Assert.AreEqual("I'm red", cellB.StringCellValue);
            Assert.AreEqual(64, cellB.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, cellB.CellStyle.FillBackgroundColor);
            Assert.AreEqual(10, cellB.CellStyle.GetFont(book).Color);
            Assert.AreEqual(0, (short)cellB.CellStyle.FillPattern);
            Assert.AreEqual("0:0:0", p.GetColor((short)64).GetHexString());
            Assert.AreEqual("FFFF:0:0", p.GetColor((short)10).GetHexString());

            // Red + green bg
            Assert.AreEqual("I'm red with a green bg", cellC.StringCellValue);
            Assert.AreEqual(11, cellC.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, cellC.CellStyle.FillBackgroundColor);
            Assert.AreEqual(10, cellC.CellStyle.GetFont(book).Color);
            Assert.AreEqual(1, (short)cellC.CellStyle.FillPattern);
            Assert.AreEqual("0:FFFF:0", p.GetColor((short)11).GetHexString());
            Assert.AreEqual("FFFF:0:0", p.GetColor((short)10).GetHexString());

            // Pink with yellow
            Assert.AreEqual("I'm pink with a yellow pattern (none)", cellD.StringCellValue);
            Assert.AreEqual(13, cellD.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, cellD.CellStyle.FillBackgroundColor);
            Assert.AreEqual(14, cellD.CellStyle.GetFont(book).Color);
            Assert.AreEqual(0, (short)cellD.CellStyle.FillPattern);
            Assert.AreEqual("FFFF:FFFF:0", p.GetColor((short)13).GetHexString());
            Assert.AreEqual("FFFF:0:FFFF", p.GetColor((short)14).GetHexString());

            // Pink with yellow - full
            Assert.AreEqual("I'm pink with a yellow pattern (full)", cellE.StringCellValue);
            Assert.AreEqual(13, cellE.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, cellE.CellStyle.FillBackgroundColor);
            Assert.AreEqual(14, cellE.CellStyle.GetFont(book).Color);
            Assert.AreEqual(0, (short)cellE.CellStyle.FillPattern);
            Assert.AreEqual("FFFF:FFFF:0", p.GetColor((short)13).GetHexString());
            Assert.AreEqual("FFFF:0:FFFF", p.GetColor((short)14).GetHexString());
        }
        [Test]
        public void TestFindSimilar()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            HSSFPalette p = book.GetCustomPalette();


            // Add a few edge colours in
            p.SetColorAtIndex((short)8, unchecked((byte)-1), (byte)0, (byte)0);
            p.SetColorAtIndex((short)9, (byte)0, unchecked((byte)-1), (byte)0);
            p.SetColorAtIndex((short)10, (byte)0, (byte)0, unchecked((byte)-1));

            // And some near a few of them
            p.SetColorAtIndex((short)11, unchecked((byte)-1), (byte)2, (byte)2);
            p.SetColorAtIndex((short)12, unchecked((byte)-2), (byte)2, (byte)10);
            p.SetColorAtIndex((short)13, unchecked((byte)-4), (byte)0, (byte)0);
            p.SetColorAtIndex((short)14, unchecked((byte)-8), (byte)0, (byte)0);

            Assert.AreEqual(
                    "FFFF:0:0", p.GetColor((short)8).GetHexString()
            );

            // Now Check we get the right stuff back
            Assert.AreEqual(
                    p.GetColor((short)8).GetHexString(),
                    p.FindSimilarColor(unchecked((byte)-1), (byte)0, (byte)0).GetHexString()
            );
            Assert.AreEqual(
                    p.GetColor((short)8).GetHexString(),
                    p.FindSimilarColor(unchecked((byte)-2), (byte)0, (byte)0).GetHexString()
            );
            Assert.AreEqual(
                    p.GetColor((short)8).GetHexString(),
                    p.FindSimilarColor(unchecked((byte)-1), (byte)1, (byte)0).GetHexString()
            );
            Assert.AreEqual(
                    p.GetColor((short)11).GetHexString(),
                    p.FindSimilarColor(unchecked((byte)-1), (byte)2, (byte)1).GetHexString()
            );
            Assert.AreEqual(
                    p.GetColor((short)12).GetHexString(),
                    p.FindSimilarColor(unchecked((byte)-1), (byte)2, (byte)10).GetHexString()
            );

            book.Close();
        }

        /**
         * Verifies that the generated gnumeric-format string values Match the
         * hardcoded values in the HSSFColor default color palette
         */
        private class ColorComparator1 : ColorComparator
        {
            public void Compare(HSSFColor expected, HSSFColor palette)
            {
                Assert.AreEqual(expected.GetHexString(), palette.GetHexString());
            }
        }

        [Test]
        public void TestGnumericStrings()
        {
            CompareToDefaults(new ColorComparator1());
        }

        /**
         * Verifies that the palette handles invalid palette indexes
         */
        private class ColorComparator2 : ColorComparator
        {
            public void Compare(HSSFColor expected, HSSFColor palette)
            {
                byte[] s1 = expected.RGB;
                byte[] s2 = palette.RGB;
                Assert.AreEqual(s1[0], s2[0]);
                Assert.AreEqual(s1[1], s2[1]);
                Assert.AreEqual(s1[2], s2[2]);
            }
        }
        [Test]
        public void TestBadIndexes()
        {
            //too small
            hssfPalette.SetColorAtIndex((short)2, (byte)255, (byte)255, (byte)255);
            //too large
            hssfPalette.SetColorAtIndex((short)0x45, (byte)255, (byte)255, (byte)255);

            //should still Match defaults; 
            CompareToDefaults(new ColorComparator2());
        }

        private void CompareToDefaults(ColorComparator c)
        {
            var colors = HSSFColor.GetIndexHash();
            IEnumerator it = colors.Keys.GetEnumerator();
            while (it.MoveNext())
            {
                int index = (int)it.Current;
                HSSFColor expectedColor = (HSSFColor)colors[index];
                HSSFColor paletteColor = hssfPalette.GetColor((short)index);
                c.Compare(expectedColor, paletteColor);
            }
        }
        [Test]
        public void TestAddColor()
        {
            try
            {
                HSSFColor hssfColor = hssfPalette.AddColor((byte)10, (byte)10, (byte)10);
                Assert.Fail();
            }
            catch (Exception)
            {
                // Failing because by default there are no colours left in the palette.
            }
        }

        private interface ColorComparator
        {
            void Compare(HSSFColor expected, HSSFColor palette);
        }
        [Test]
        public void Test48403()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            var color = new Rgb24(0, 0x6B, 0x6B); //decode("#006B6B");
            HSSFPalette palette = wb.GetCustomPalette();

            HSSFColor hssfColor = palette.FindColor(color.R, color.G, color.B);
            Assert.IsNull(hssfColor);

            palette.SetColorAtIndex(
                    (short)(PaletteRecord.STANDARD_PALETTE_SIZE - 1),
                    (byte)color.R, (byte)color.G,
                    (byte)color.B);
            hssfColor = palette.GetColor((short)(PaletteRecord.STANDARD_PALETTE_SIZE - 1));
            Assert.IsNotNull(hssfColor);
            Assert.AreEqual(55, hssfColor.Indexed);
            CollectionAssert.AreEqual(new short[] { 0, 107, 107 }, hssfColor.GetTriplet());

            wb.Close();
        }
    }
}
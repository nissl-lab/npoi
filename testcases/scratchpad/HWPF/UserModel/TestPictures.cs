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


namespace TestCases.HWPF.UserModel
{
    using NPOI.HWPF.UserModel;
    
    using System.Collections.Generic;
    using NPOI.HWPF;
    using NPOI.HWPF.Model;
    using NPOI.Util;
    using NUnit.Framework;
    /**
     * Test the picture handling
     *
     * @author Nick Burch
     */
    [TestFixture]
    public class TestPictures
    {

        /**
         * two jpegs
         */
        [Test]
        public void TestTwoImages()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("two_images.doc");
            List<Picture> pics = doc.GetPicturesTable().GetAllPictures();

            Assert.IsNotNull(pics);
            Assert.AreEqual(pics.Count, 2);
            for (int i = 0; i < pics.Count; i++)
            {
                Picture pic = (Picture)pics[i];
                Assert.IsNotNull(pic.SuggestFileExtension());
                Assert.IsNotNull(pic.SuggestFullFileName());
            }

            Picture picA = pics[0];
            Picture picB = pics[1];
            Assert.AreEqual("jpg", picA.SuggestFileExtension());
            Assert.AreEqual("jpg", picA.SuggestFileExtension());
        }

        /**
         * pngs and jpegs
         */
        [Test]
        public void TestDifferentImages()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("testPictures.doc");
            List<Picture> pics = doc.GetPicturesTable().GetAllPictures();

            Assert.IsNotNull(pics);
            Assert.AreEqual(7, pics.Count);
            for (int i = 0; i < pics.Count; i++)
            {
                Picture pic = (Picture)pics[i];
                Assert.IsNotNull(pic.SuggestFileExtension());
                Assert.IsNotNull(pic.SuggestFullFileName());
            }

            Assert.AreEqual("jpg", pics[0].SuggestFileExtension());
            Assert.AreEqual("image/jpeg", pics[0].MimeType);
            Assert.AreEqual("jpg", pics[1].SuggestFileExtension());
            Assert.AreEqual("image/jpeg", pics[1].MimeType);
            Assert.AreEqual("png", pics[3].SuggestFileExtension());
            Assert.AreEqual("image/png", pics[3].MimeType);
            Assert.AreEqual("png", pics[4].SuggestFileExtension());
            Assert.AreEqual("image/png", pics[4].MimeType);
            Assert.AreEqual("wmf", pics[5].SuggestFileExtension());
            Assert.AreEqual("image/x-wmf", pics[5].MimeType);
            Assert.AreEqual("jpg", pics[6].SuggestFileExtension());
            Assert.AreEqual("image/jpeg", pics[6].MimeType);
        }

        /**
         * emf image, nice and simple
         */
        [Test]
        public void TestEmfImage()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("vector_image.doc");
            List<Picture> pics = doc.GetPicturesTable().GetAllPictures();

            Assert.IsNotNull(pics);
            Assert.AreEqual(1, pics.Count);

            Picture pic = pics[0];
            Assert.IsNotNull(pic.SuggestFileExtension());
            Assert.IsNotNull(pic.SuggestFullFileName());
            Assert.IsTrue(pic.Size > 128);

            // Check right contents
            byte[] emf = POIDataSamples.GetDocumentInstance().ReadFile("vector_image.emf");
            byte[] pemf = pic.GetContent();
            Assert.AreEqual(emf.Length, pemf.Length);
            for (int i = 0; i < emf.Length; i++)
            {
                Assert.AreEqual(emf[i], pemf[i]);
            }
        }

        /**
         * emf image, with a crazy offset
         */
        [Test]
        public void TestEmfComplexImage()
        {

            // Commenting out this Test case temporarily. The file emf_2003_image does not contain any
            // pictures. Instead it has an office drawing object. Need to rewrite this Test after
            // revisiting the implementation of office drawing objects.

            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug41898.doc");
            List<Picture> pics = doc.GetPicturesTable().GetAllPictures();

            Assert.IsNotNull(pics);
            Assert.AreEqual(1, pics.Count);

            Picture pic = pics[0];
            Assert.IsNotNull(pic.SuggestFileExtension());
            Assert.IsNotNull(pic.SuggestFullFileName());

            // This one's tricky
            // TODO: Fix once we've sorted bug #41898
            Assert.IsNotNull(pic.GetContent());
            Assert.IsNotNull(pic.GetRawContent());

            // These are probably some sort of offSet, need to figure them out
            Assert.AreEqual(4, pic.Size);
            Assert.AreEqual((uint)0x80000000, LittleEndian.GetUInt(pic.GetContent()));
            Assert.AreEqual((uint)0x80000000, LittleEndian.GetUInt(pic.GetRawContent()));
        }
        [Test]
        public void TestPicturesWithTable()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug44603.doc");

            List<Picture> pics = doc.GetPicturesTable().GetAllPictures();
            Assert.AreEqual(2, pics.Count);
        }
        [Test]
        public void TestPicturesInHeader()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("header_image.doc");

            List<Picture> pics = doc.GetPicturesTable().GetAllPictures();
            Assert.AreEqual(2, pics.Count);
        }
        [Test]
        public void TestFastSaved()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("rasp.doc");

            doc.GetPicturesTable().GetAllPictures(); // just check that we do not throw Exception
        }
        [Test]
        public void TestFastSaved2()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("o_kurs.doc");

            doc.GetPicturesTable().GetAllPictures(); // just check that we do not throw Exception
        }
        [Test]
        public void TestFastSaved3()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("ob_is.doc");

            doc.GetPicturesTable().GetAllPictures(); // just check that we do not throw Exception
        }

        /**
         * When you embed another office document into Word, it stores
         *  a rendered "icon" picture of what that document looks like.
         * This image is re-Created when you edit the embeded document,
         *  then used as-is to speed things up.
         * Check that we can properly read one of these
         */
        [Test]
        public void TestEmbededDocumentIcon()
        {
            // This file has two embeded excel files, an embeded powerpoint
            //   file and an embeded word file, in that order
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("word_with_embeded.doc");

            // Check we don't break loading the pictures
            doc.GetPicturesTable().GetAllPictures();
            PicturesTable pictureTable = doc.GetPicturesTable();

            // Check the Text, and its embeded images
            Paragraph p;
            Range r = doc.GetRange();
            Assert.AreEqual(1, r.NumSections);
            Assert.AreEqual(5, r.NumParagraphs);

            p = r.GetParagraph(0);
            Assert.AreEqual(2, p.NumCharacterRuns);
            Assert.AreEqual("I have lots of embedded files in me\r", p.Text);
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(0)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(1)));

            p = r.GetParagraph(1);
            Assert.AreEqual(5, p.NumCharacterRuns);
            Assert.AreEqual("\u0013 EMBED Excel.Sheet.8  \u0014\u0001\u0015\r", p.Text);
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(0)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(1)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(2)));
            Assert.AreEqual(true, pictureTable.HasPicture(p.GetCharacterRun(3)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(4)));

            p = r.GetParagraph(2);
            Assert.AreEqual(6, p.NumCharacterRuns);
            Assert.AreEqual("\u0013 EMBED Excel.Sheet.8  \u0014\u0001\u0015\r", p.Text);
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(0)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(1)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(2)));
            Assert.AreEqual(true, pictureTable.HasPicture(p.GetCharacterRun(3)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(4)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(5)));

            p = r.GetParagraph(3);
            Assert.AreEqual(6, p.NumCharacterRuns);
            Assert.AreEqual("\u0013 EMBED PowerPoint.Show.8  \u0014\u0001\u0015\r", p.Text);
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(0)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(1)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(2)));
            Assert.AreEqual(true, pictureTable.HasPicture(p.GetCharacterRun(3)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(4)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(5)));

            p = r.GetParagraph(4);
            Assert.AreEqual(6, p.NumCharacterRuns);
            Assert.AreEqual("\u0013 EMBED Word.Document.8 \\s \u0014\u0001\u0015\r", p.Text);
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(0)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(1)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(2)));
            Assert.AreEqual(true, pictureTable.HasPicture(p.GetCharacterRun(3)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(4)));
            Assert.AreEqual(false, pictureTable.HasPicture(p.GetCharacterRun(5)));

            // Look at the pictures table
            List<Picture> pictures = pictureTable.GetAllPictures();
            Assert.AreEqual(4, pictures.Count);

            Picture picture = pictures[0];
            Assert.AreEqual("", picture.SuggestFileExtension());
            Assert.AreEqual("0", picture.SuggestFullFileName());
            Assert.AreEqual("image/unknown", picture.MimeType);

            picture = pictures[1];
            Assert.AreEqual("", picture.SuggestFileExtension());
            Assert.AreEqual("469", picture.SuggestFullFileName());
            Assert.AreEqual("image/unknown", picture.MimeType);

            picture = pictures[2];
            Assert.AreEqual("", picture.SuggestFileExtension());
            Assert.AreEqual("8c7", picture.SuggestFullFileName());
            Assert.AreEqual("image/unknown", picture.MimeType);

            picture = pictures[3];
            Assert.AreEqual("", picture.SuggestFileExtension());
            Assert.AreEqual("10a8", picture.SuggestFullFileName());
            Assert.AreEqual("image/unknown", picture.MimeType);
        }
        [Test]
        public void TestEquation()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("equation.doc");
            PicturesTable pictures = doc.GetPicturesTable();

            List<Picture> allPictures = pictures.GetAllPictures();
            Assert.AreEqual(1, allPictures.Count);

            Picture picture = allPictures[0];
            Assert.IsNotNull(picture);
            Assert.AreEqual(PictureType.EMF, picture.SuggestPictureType());
            Assert.AreEqual(PictureType.EMF.Extension,
                    picture.SuggestFileExtension());
            Assert.AreEqual(PictureType.EMF.Mime, picture.MimeType);
            Assert.AreEqual("0.emf", picture.SuggestFullFileName());
        }

        /**
         * In word you can have floating or fixed pictures.
         * Fixed have a \u0001 in place with an offset to the
         *  picture data.
         * Floating have a \u0008 in place, which references a
         *  \u0001 which has the offSet. More than one can
         *  reference the same \u0001
         */
        [Test]
        public void TestFloatingPictures()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("FloatingPictures.doc");
            PicturesTable pictures = doc.GetPicturesTable();

            // There are 19 images in the picture, but some are
            //  duplicate floating ones
            Assert.AreEqual(17, pictures.GetAllPictures().Count);

            int plain8s = 0;
            int escher8s = 0;
            int image1s = 0;

            Range r = doc.GetRange();
            for (int np = 0; np < r.NumParagraphs; np++)
            {
                Paragraph p = r.GetParagraph(np);
                for (int nc = 0; nc < p.NumCharacterRuns; nc++)
                {
                    CharacterRun cr = p.GetCharacterRun(nc);
                    if (pictures.HasPicture(cr))
                    {
                        image1s++;
                    }
                    else if (pictures.HasEscherPicture(cr))
                    {
                        escher8s++;
                    }
                    else if (cr.Text.StartsWith("\u0008"))
                    {
                        plain8s++;
                    }
                }
            }
            // Total is 20, as the 4 escher 8s all reference
            //  the same regular image
            Assert.AreEqual(16, image1s);
            Assert.AreEqual(4, escher8s);
            Assert.AreEqual(0, plain8s);
        }
        [Test]
        public void TestCroppedPictures()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("testCroppedPictures.doc");
            List<Picture> pics = doc.GetPicturesTable().GetAllPictures();

            Assert.IsNotNull(pics);
            Assert.AreEqual(2, pics.Count);

            Picture pic1 = pics[0];
            Assert.AreEqual(27, pic1.AspectRatioX);
            Assert.AreEqual(270, pic1.HorizontalScalingFactor);
            Assert.AreEqual(27, pic1.AspectRatioY);
            Assert.AreEqual(271, pic1.VerticalScalingFactor);
            Assert.AreEqual(12000, pic1.DxaGoal);       // 21.17 cm / 2.54 cm/inch * 72dpi * 20 = 12000
            Assert.AreEqual(9000, pic1.DyaGoal);        // 15.88 cm / 2.54 cm/inch * 72dpi * 20 = 9000
            Assert.AreEqual(0, pic1.DxaCropLeft);
            Assert.AreEqual(0, pic1.DxaCropRight);
            Assert.AreEqual(0, pic1.DyaCropTop);
            Assert.AreEqual(0, pic1.DyaCropBottom);

            Picture pic2 = pics[1];
            Assert.AreEqual(76, pic2.AspectRatioX);
            Assert.AreEqual(764, pic2.HorizontalScalingFactor);
            Assert.AreEqual(68, pic2.AspectRatioY);
            Assert.AreEqual(685, pic2.VerticalScalingFactor);
            Assert.AreEqual(12000, pic2.DxaGoal);       // 21.17 cm / 2.54 cm/inch * 72dpi * 20 = 12000
            Assert.AreEqual(9000, pic2.DyaGoal);        // 15.88 cm / 2.54 cm/inch * 72dpi * 20 = 9000
            Assert.AreEqual(0, pic2.DxaCropLeft);       // TODO YK: The Picture is cropped but HWPF reads the crop parameters all zeros
            Assert.AreEqual(0, pic2.DxaCropRight);
            Assert.AreEqual(0, pic2.DyaCropTop);
            Assert.AreEqual(0, pic2.DyaCropBottom);
        }
    }
}
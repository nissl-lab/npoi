namespace TestCases.XWPF.UserModel
{
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using NPOI.XWPF.Model;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using System.IO;

    [TestFixture]
    public class TestXWPFComment
    {
        [Test]
        public void TestText()
        {
            using (XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("comment.docx"))
            {
                Assert.AreEqual(1, doc.GetComments().Length);
                XWPFComment comment = doc.GetComments()[0];
                Assert.AreEqual("Unbekannter Autor", comment.GetAuthor());
                Assert.AreEqual("0", comment.GetId());
                Assert.AreEqual("This is the first line\n\nThis is the second line", comment.GetText());
            }
        }

        [Test]
        public void TestAddComment()
        {
            var cId = 0;
            var date = LocaleUtil.GetLocaleCalendar().ToString();
            using (XWPFDocument docOut = new XWPFDocument())
            {
                Assert.IsNull(docOut.GetDocComments());

                XWPFComments comments = docOut.CreateComments();
                XWPFComment comment = comments.CreateComment(cId.ToString());
                comment.Author = "Author";
                comment.Initials = "s";
                comment.Date = date;

                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);
                Assert.AreEqual(1, docIn.GetComments().Length);
                comment = docIn.GetCommentByID(cId.ToString());
                Assert.IsNotNull(comment);
                Assert.AreEqual("Author", comment.GetAuthor());
                Assert.AreEqual("s", comment.GetInitials());
                //Assert.AreEqual(date.getTimeInMillis(), comment.getDate().getTimeInMillis());
                Assert.AreEqual(date, comment.GetDate());
            }
        }

        [Test]
        public void TestRemoveComment()
        {
            using (XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("comment.docx"))
            {
                Assert.AreEqual(1, doc.GetComments().Length);

                doc.GetDocComments().RemoveComment(0);

                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(doc);
                Assert.AreEqual(0, docIn.GetComments().Length);
            }
        }

        [Test]
        public void TestCreateParagraph()
        {
            using (XWPFDocument doc = new XWPFDocument())
            {
                XWPFComments comments = doc.CreateComments();
                XWPFComment comment = comments.CreateComment("1");
                XWPFParagraph paragraph = comment.CreateParagraph();
                paragraph.CreateRun().SetText("comment paragraph text");

                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(doc);
                XWPFComment xwpfComment = docIn.GetCommentByID("1");
                Assert.AreEqual(1, xwpfComment.Paragraphs.Count);
                String text = xwpfComment.GetParagraphArray(0).Text;
                Assert.AreEqual("comment paragraph text", text);
            }
        }

        [Test]
        public void TestAddPicture()
        {
            using (XWPFDocument doc = new XWPFDocument())
            {
                XWPFComments comments = doc.CreateComments();
                XWPFComment comment = comments.CreateComment("1");
                XWPFParagraph paragraph = comment.CreateParagraph();
                XWPFRun r = paragraph.CreateRun();
                r.AddPicture(new ByteArrayInputStream(new byte[0]),
                        (int)PictureType.JPEG/*Document.PICTURE_TYPE_JPEG*/, "test.jpg", 21, 32);

                Assert.AreEqual(1, comments.GetAllPictures().Count);
                Assert.AreEqual(1, doc.AllPackagePictures.Count);
            }
        }

        [Test]
        public void TestCreateTable()
        {
            using (XWPFDocument doc = new XWPFDocument())
            {
                XWPFComments comments = doc.CreateComments();
                XWPFComment comment = comments.CreateComment("1");
                comment.CreateTable(1, 1);

                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(doc);
                XWPFComment xwpfComment = docIn.GetCommentByID("1");
                Assert.AreEqual(1, xwpfComment.Tables.Count);
            }
        }
    }
}

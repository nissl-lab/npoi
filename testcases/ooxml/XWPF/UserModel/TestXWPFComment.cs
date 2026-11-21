namespace TestCases.XWPF.UserModel
{
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using NPOI.XWPF.Model;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
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
                ClassicAssert.AreEqual(1, doc.GetComments().Length);
                XWPFComment comment = doc.GetComments()[0];
                ClassicAssert.AreEqual("Unbekannter Autor", comment.GetAuthor());
                ClassicAssert.AreEqual("0", comment.GetId());
                ClassicAssert.AreEqual("This is the first line\n\nThis is the second line", comment.GetText());
            }
        }

        [Test]
        public void TestAddComment()
        {
            var cId = 0;
            var date = LocaleUtil.GetLocaleCalendar().ToString();
            using (XWPFDocument docOut = new XWPFDocument())
            {
                ClassicAssert.IsNull(docOut.GetDocComments());

                XWPFComments comments = docOut.CreateComments();
                XWPFComment comment = comments.CreateComment(cId.ToString());
                comment.Author = "Author";
                comment.Initials = "s";
                comment.Date = date;

                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);
                ClassicAssert.AreEqual(1, docIn.GetComments().Length);
                comment = docIn.GetCommentByID(cId.ToString());
                ClassicAssert.IsNotNull(comment);
                ClassicAssert.AreEqual("Author", comment.GetAuthor());
                ClassicAssert.AreEqual("s", comment.GetInitials());
                //ClassicAssert.AreEqual(date.getTimeInMillis(), comment.getDate().getTimeInMillis());
                ClassicAssert.AreEqual(date, comment.GetDate());
            }
        }

        [Test]
        public void TestRemoveComment()
        {
            using (XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("comment.docx"))
            {
                ClassicAssert.AreEqual(1, doc.GetComments().Length);

                doc.GetDocComments().RemoveComment(0);

                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(doc);
                ClassicAssert.AreEqual(0, docIn.GetComments().Length);
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
                ClassicAssert.AreEqual(1, xwpfComment.Paragraphs.Count);
                String text = xwpfComment.GetParagraphArray(0).Text;
                ClassicAssert.AreEqual("comment paragraph text", text);
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

                ClassicAssert.AreEqual(1, comments.GetAllPictures().Count);
                ClassicAssert.AreEqual(1, doc.AllPackagePictures.Count);
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
                ClassicAssert.AreEqual(1, xwpfComment.Tables.Count);
            }
        }
    }
}

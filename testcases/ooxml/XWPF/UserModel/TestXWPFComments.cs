namespace TestCases.XWPF.UserModel
{
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection.Metadata;

    [TestFixture]
    public class TestXWPFComments
    {
        [Test]
        public void TestAddCommentsToDoc()
        {
            var cId = "0";
            using (XWPFDocument docOut = new XWPFDocument())
            {
                ClassicAssert.IsNull(docOut.GetDocComments());

                // create comments
                XWPFComments comments = docOut.CreateComments();
                ClassicAssert.IsNotNull(comments);
                ClassicAssert.AreSame(comments, docOut.CreateComments());

                // create comment
                XWPFComment comment = comments.CreateComment(cId);
                comment.Author = "Author";
                comment.CreateParagraph().CreateRun().SetText("comment paragraph");

                // apply comment to run text
                XWPFParagraph paragraph = docOut.CreateParagraph();
                paragraph.GetCTP().AddNewCommentRangeStart().id = cId;
                paragraph.GetCTP().AddNewR().AddNewT().Value = "HelloWorld";
                paragraph.GetCTP().AddNewCommentRangeEnd().id = cId;
                paragraph.GetCTP().AddNewR().AddNewCommentReference().id = cId;

                // check
                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);
                ClassicAssert.IsNotNull(docIn.GetDocComments());
                ClassicAssert.AreEqual(1, docIn.GetComments().Length);
                comment = docIn.GetCommentByID("0");
                ClassicAssert.IsNotNull(comment);
                ClassicAssert.AreEqual("Author", comment.GetAuthor());
            }
        }

        [Test]
        public void TestReadComments()
        {
            using (XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("testComment.docx"))
            {
                XWPFComments docComments = doc.GetDocComments();
                ClassicAssert.IsNotNull(docComments);
                XWPFComment[] comments = doc.GetComments();
                ClassicAssert.AreEqual(1, comments.Length);

                IList<XWPFPictureData> allPictures = docComments.GetAllPictures();
                ClassicAssert.AreEqual(1, allPictures.Count);
            }
        }

        [Test]
        public void TestNPOIBug1481()
        {
            using(XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("NPOI-bug-1481.docx"))
            {
                XWPFComment[] comments = doc.GetComments();
                ClassicAssert.AreEqual(1, comments.Length);

                XWPFComment comment = comments[0];
                ClassicAssert.AreEqual("Claudio Pais", comment.GetAuthor());
                ClassicAssert.AreEqual("2025-01-23T17:18:00Z", comment.Date);
                ClassicAssert.AreEqual("Bla bla", comment.GetText());

                XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(doc);

                comments = docIn.GetComments();
                ClassicAssert.AreEqual(1, comments.Length);

                comment = comments[0];
                ClassicAssert.AreEqual("Claudio Pais", comment.GetAuthor());
                ClassicAssert.AreEqual("2025-01-23T17:18:00Z", comment.Date);
                ClassicAssert.AreEqual("Bla bla", comment.GetText());

            }
        }
    }
}

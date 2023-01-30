namespace TestCases.XWPF.UserModel
{
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using System.Collections.Generic;

    [TestFixture]
    public class TestXWPFComments
    {
        [Test]
        public void TestAddCommentsToDoc()
        {
            var cId = "0";
            using (XWPFDocument docOut = new XWPFDocument())
            {
                Assert.IsNull(docOut.GetDocComments());

                // create comments
                XWPFComments comments = docOut.CreateComments();
                Assert.IsNotNull(comments);
                Assert.AreSame(comments, docOut.CreateComments());

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
                Assert.IsNotNull(docIn.GetDocComments());
                Assert.AreEqual(1, docIn.GetComments().Length);
                comment = docIn.GetCommentByID("0");
                Assert.IsNotNull(comment);
                Assert.AreEqual("Author", comment.GetAuthor());
            }
        }

        [Test]
        public void TestReadComments()
        {
            using (XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("testComment.docx"))
            {
                XWPFComments docComments = doc.GetDocComments();
                Assert.IsNotNull(docComments);
                XWPFComment[] comments = doc.GetComments();
                Assert.AreEqual(1, comments.Length);

                IList<XWPFPictureData> allPictures = docComments.GetAllPictures();
                Assert.AreEqual(1, allPictures.Count);
            }
        }
    }
}

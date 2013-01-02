using System.IO;
using NUnit.Framework;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace ooxml.Testcases
{
    /// <summary>
    ///This is a test class for CommentsDocumentTest and is intended
    ///to contain all CommentsDocumentTest Unit Tests
    ///</summary>
    [TestFixture]
    public class CommentsDocumentTest
    {
        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


//        /// <summary>
//        ///A test for the Serialization of CT_Comments.
//        ///</summary>
//        [Test]
//        public void SerializeCommentsDocumentTest()
//        {
//            CT_Comments comments = new CT_Comments();
//            comments.AddNewAuthors().AddAuthor("Christian");
//            comments.authors.AddAuthor("Tony");
//            comments.AddNewCommentList();

//            CT_Comment singleComment = comments.commentList.AddNewComment();
//            singleComment.authorId = 1;
//            singleComment.@ref = "A7";
//            var text = new CT_Rst();
//            text.r = new System.Collections.Generic.List<CT_RElt>();
//            var commentText = new CT_RElt();
//            commentText.t = "First Comment";
//            text.r.Add(commentText);
//            singleComment.text = text;

//            using (StringWriter stream = new StringWriter())
//            {
//                CommentsDocument_Accessor.serializer.Serialize(stream, comments);//, CommentsDocument_Accessor.namespaces);
//                string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
//<comments xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
//  <authors>
//    <author>Christian</author>
//    <author>Tony</author>
//  </authors>
//  <commentList>
//    <comment ref=""A7"" authorId=""1"">
//      <text>
//        <r>
//          <t>First Comment</t>
//        </r>
//      </text>
//    </comment>
//  </commentList>
//</comments>";
//                Assert.AreEqual(expected, stream.ToString());
//            }
//        }


//        /// <summary>
//        ///A test for Deserialize
//        ///</summary>
//        [Test]
//        public void DeserializeCommentsDocumentTest()
//        {
//            // The following is the excerpt of an Excel file.
//            string input =
//@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
//<comments xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
//   <authors>
//      <author>First Author</author>
//   </authors>
//   <commentList>
//      <comment ref=""A3"" authorId=""0"">
//         <text>
//            <r>
//               <rPr>
//                  <b/>
//                  <sz val=""9""/>
//                  <color indexed=""81""/>
//                  <rFont val=""Tahoma""/>
//                  <charset val=""1""/>
//               </rPr>
//               <t>Autor:</t>
//            </r>
//            <r>
//               <rPr>
//                  <sz val=""9""/>
//                  <color indexed=""81""/>
//                  <rFont val=""Tahoma""/>
//                  <charset val=""1""/>
//               </rPr>
//               <t xml:space=""preserve"">
//First Column on the first sheet; This is just an increasing column</t>
//            </r>
//         </text>
//      </comment>
//      <comment ref=""B3"" authorId=""0"">
//         <text>
//            <r>
//               <rPr>
//                  <b/>
//                  <sz val=""9""/>
//                  <color indexed=""81""/>
//                  <rFont val=""Tahoma""/>
//                  <charset val=""1""/>
//               </rPr>
//               <t>Autor:</t>
//            </r>
//            <r>
//               <rPr>
//                  <sz val=""9""/>
//                  <color indexed=""81""/>
//                  <rFont val=""Tahoma""/>
//                  <charset val=""1""/>
//               </rPr>
//               <t xml:space=""preserve"">
//This column multiplies the first column value with 2.</t>
//            </r>
//         </text>
//      </comment>
//      <comment ref=""C3"" authorId=""0"">
//         <text>
//            <r>
//               <rPr>
//                  <b/>
//                  <sz val=""9""/>
//                  <color indexed=""81""/>
//                  <rFont val=""Tahoma""/>
//                  <charset val=""1""/>
//               </rPr>
//               <t>Autor:</t>
//            </r>
//            <r>
//               <rPr>
//                  <sz val=""9""/>
//                  <color indexed=""81""/>
//                  <rFont val=""Tahoma""/>
//                  <charset val=""1""/>
//               </rPr>
//               <t xml:space=""preserve"">
//This column adds the first and second one.</t>
//            </r>
//         </text>
//      </comment>
//   </commentList>
//</comments>";
//            CT_Comments result;
//            //{
//            //    StringReader stream = new StringReader(input);
//            //    result = (CT_Comments)CommentsDocument_Accessor.serializer.Deserialize(stream); // instantiate source code to enable debugging the serialization code
//            //}
//            {
//                using (StringReader stream = new StringReader(input))
//                {
//                    result = (CT_Comments)CommentsDocument_Accessor.serializer.Deserialize(stream);
//                }
//            }
//            Assert.AreEqual(3, result.commentList.SizeOfCommentArray());
//            Assert.AreEqual(1, result.authors.SizeOfAuthorArray());
//            Assert.AreEqual("First Author", result.authors.author[0]);

//            // First Comment
//            //<comment ref=""A3"" authorId=""0"">
//            //  <text>
//            //    <r>
//            //        <rPr>
//            //            <b/>
//            //            <sz val=""9""/>
//            //            <color indexed=""81""/>
//            //            <rFont val=""Tahoma""/>
//            //            <charset val=""1""/>
//            //        </rPr>
//            //        <t>Autor:</t>
//            //    </r>
//            //    <r>
//            //        <rPr>
//            //            <sz val=""9""/>
//            //            <color indexed=""81""/>
//            //            <rFont val=""Tahoma""/>
//            //            <charset val=""1""/>
//            //        </rPr>
//            //        <t xml:space=""preserve"">
//            //First Column on the first sheet; This is just an increasing column</t>
//            //    </r>
//            //  </text>
//            //</comment>
//            {
//                var comment = result.commentList.GetCommentArray(0);
//                Assert.AreEqual("A3", comment.@ref);
//                Assert.AreEqual(0u, comment.authorId);
//                Assert.AreEqual("Autor:", comment.text.r[0].t);
//                Assert.AreEqual("\nFirst Column on the first sheet; This is just an increasing column", comment.text.r[1].t);

//                // TODO fix Font extraction Assert.AreEqual("Tahoma", comment.text.r[0].rPr.GetRFontArray(0).val);
//            }

//            // Second Comment
//            {
//                var comment = result.commentList.GetCommentArray(1);
//                Assert.AreEqual("B3", comment.@ref);
//                Assert.AreEqual(0u, comment.authorId);
//                Assert.AreEqual("Autor:", comment.text.r[0].t);
//                Assert.AreEqual("This column multiplies the first column value with 2.", comment.text.r[1].t.Trim());

//                // TODO Assert.AreEqual("Tahoma", comment.text.r[0].rPr.GetRFontArray(0).val);
//            }

//            // Third Comment
//            {
//                var comment = result.commentList.GetCommentArray(2);
//                Assert.AreEqual("C3", comment.@ref);
//                Assert.AreEqual(0u, comment.authorId);
//                Assert.AreEqual("Autor:", comment.text.r[0].t);
//                Assert.AreEqual("\nThis column adds the first and second one.", comment.text.r[1].t);

//               // TODO Assert.AreEqual("Tahoma", comment.text.r[0].rPr.GetRFontArray(0).val);
//            }
//        }



    }
}

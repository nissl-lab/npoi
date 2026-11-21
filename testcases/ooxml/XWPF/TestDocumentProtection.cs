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
namespace TestCases.XWPF
{
    using System.IO;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.POIFS.Crypt;
    using NPOI.Util;
    using NPOI.XWPF.UserModel;

    [TestFixture]
    public class TestDocumentProtection
    {
        [Test]
        public void TestShouldReadEnforcementProperties()
        {
            XWPFDocument documentWithoutDocumentProtectionTag = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            ClassicAssert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedReadonlyProtection());
            ClassicAssert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedFillingFormsProtection());
            ClassicAssert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedCommentsProtection());
            ClassicAssert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedTrackedChangesProtection());
            documentWithoutDocumentProtectionTag.Close();

            XWPFDocument documentWithoutEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection_tag_existing.docx");
            ClassicAssert.IsFalse(documentWithoutEnforcement.IsEnforcedReadonlyProtection());
            ClassicAssert.IsFalse(documentWithoutEnforcement.IsEnforcedFillingFormsProtection());
            ClassicAssert.IsFalse(documentWithoutEnforcement.IsEnforcedCommentsProtection());
            ClassicAssert.IsFalse(documentWithoutEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithoutEnforcement.Close();

            XWPFDocument documentWithReadonlyEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_readonly_no_password.docx");
            ClassicAssert.IsTrue(documentWithReadonlyEnforcement.IsEnforcedReadonlyProtection());
            ClassicAssert.IsFalse(documentWithReadonlyEnforcement.IsEnforcedFillingFormsProtection());
            ClassicAssert.IsFalse(documentWithReadonlyEnforcement.IsEnforcedCommentsProtection());
            ClassicAssert.IsFalse(documentWithReadonlyEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithReadonlyEnforcement.Close();

            XWPFDocument documentWithFillingFormsEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_forms_no_password.docx");
            ClassicAssert.IsTrue(documentWithFillingFormsEnforcement.IsEnforcedFillingFormsProtection());
            ClassicAssert.IsFalse(documentWithFillingFormsEnforcement.IsEnforcedReadonlyProtection());
            ClassicAssert.IsFalse(documentWithFillingFormsEnforcement.IsEnforcedCommentsProtection());
            ClassicAssert.IsFalse(documentWithFillingFormsEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithFillingFormsEnforcement.Close();

            XWPFDocument documentWithCommentsEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_comments_no_password.docx");
            ClassicAssert.IsFalse(documentWithCommentsEnforcement.IsEnforcedFillingFormsProtection());
            ClassicAssert.IsFalse(documentWithCommentsEnforcement.IsEnforcedReadonlyProtection());
            ClassicAssert.IsTrue(documentWithCommentsEnforcement.IsEnforcedCommentsProtection());
            ClassicAssert.IsFalse(documentWithCommentsEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithCommentsEnforcement.Close();

            XWPFDocument documentWithTrackedChangesEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_trackedChanges_no_password.docx");
            ClassicAssert.IsFalse(documentWithTrackedChangesEnforcement.IsEnforcedFillingFormsProtection());
            ClassicAssert.IsFalse(documentWithTrackedChangesEnforcement.IsEnforcedReadonlyProtection());
            ClassicAssert.IsFalse(documentWithTrackedChangesEnforcement.IsEnforcedCommentsProtection());
            ClassicAssert.IsTrue(documentWithTrackedChangesEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithTrackedChangesEnforcement.Close();
        }

        [Test]
        public void TestShouldEnforceForReadOnly()
        {
            //		XWPFDocument document = CreateDocumentFromSampleFile("test-data/document/documentProtection_no_protection.docx");
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            ClassicAssert.IsFalse(document.IsEnforcedReadonlyProtection());

            document.EnforceReadonlyProtection();

            ClassicAssert.IsTrue(document.IsEnforcedReadonlyProtection());
            document.Close();
        }

        [Test]
        public void TestShouldEnforceForFillingForms()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            ClassicAssert.IsFalse(document.IsEnforcedFillingFormsProtection());

            document.EnforceFillingFormsProtection();

            ClassicAssert.IsTrue(document.IsEnforcedFillingFormsProtection());
            document.Close();
        }

        [Test]
        public void TestShouldEnforceForComments()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            ClassicAssert.IsFalse(document.IsEnforcedCommentsProtection());

            document.EnforceCommentsProtection();

            ClassicAssert.IsTrue(document.IsEnforcedCommentsProtection());
            document.Close();
        }

        [Test]
        public void TestShouldEnforceForTrackedChanges()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            ClassicAssert.IsFalse(document.IsEnforcedTrackedChangesProtection());

            document.EnforceTrackedChangesProtection();

            ClassicAssert.IsTrue(document.IsEnforcedTrackedChangesProtection());
            document.Close();
        }

        [Test]
        public void TestShouldUnsetEnforcement()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_readonly_no_password.docx");
            ClassicAssert.IsTrue(document.IsEnforcedReadonlyProtection());

            document.RemoveProtectionEnforcement();

            ClassicAssert.IsFalse(document.IsEnforcedReadonlyProtection());
            document.Close();
        }

        [Test]
        public void TestIntegration()
        {
            XWPFDocument doc1 = new XWPFDocument();

            XWPFParagraph p1 = doc1.CreateParagraph();

            XWPFRun r1 = p1.CreateRun();
            r1.SetText("Lorem ipsum dolor sit amet.");
            doc1.EnforceCommentsProtection();

            FileInfo tempFile = TempFile.CreateTempFile("documentProtectionFile", ".docx");
            if (File.Exists(tempFile.FullName))
                File.Delete(tempFile.FullName);
            Stream out1 = new FileStream(tempFile.FullName, FileMode.CreateNew);

            doc1.Write(out1);
            out1.Close();

            FileStream inputStream = new FileStream(tempFile.FullName, FileMode.Open);
            XWPFDocument doc2 = new XWPFDocument(inputStream);
            inputStream.Close();

            ClassicAssert.IsTrue(doc2.IsEnforcedCommentsProtection());
            doc2.Close();
            doc1.Close();
        }

        [Test]
        public void TestUpdateFields()
        {
            XWPFDocument doc = new XWPFDocument();
            ClassicAssert.IsFalse(doc.IsEnforcedUpdateFields());
            doc.EnforceUpdateFields();
            ClassicAssert.IsTrue(doc.IsEnforcedUpdateFields());
            doc.Close();
        }

        [Test]
        public void Bug56076_read()
        {
            // test legacy xored-hashed password
            ClassicAssert.AreEqual("64CEED7E", CryptoFunctions.XorHashPassword("Example"));
            // check leading 0
            ClassicAssert.AreEqual("0005CB00", CryptoFunctions.XorHashPassword("34579"));

            // test document write protection with password
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("bug56076.docx");
            bool isValid = document.ValidateProtectionPassword("Example");
            ClassicAssert.IsTrue(isValid);
        }

        [Test]
        public void Bug56076_write()
        {
            // test document write protection with password
            XWPFDocument document = new XWPFDocument();
            document.EnforceCommentsProtection("Example", HashAlgorithm.sha512);
            document = XWPFTestDataSamples.WriteOutAndReadBack(document);
            bool isValid = document.ValidateProtectionPassword("Example");
            ClassicAssert.IsTrue(isValid);
        }
    }
}

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
    using NUnit.Framework;
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
            Assert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedReadonlyProtection());
            Assert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedFillingFormsProtection());
            Assert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedCommentsProtection());
            Assert.IsFalse(documentWithoutDocumentProtectionTag.IsEnforcedTrackedChangesProtection());
            documentWithoutDocumentProtectionTag.Close();

            XWPFDocument documentWithoutEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection_tag_existing.docx");
            Assert.IsFalse(documentWithoutEnforcement.IsEnforcedReadonlyProtection());
            Assert.IsFalse(documentWithoutEnforcement.IsEnforcedFillingFormsProtection());
            Assert.IsFalse(documentWithoutEnforcement.IsEnforcedCommentsProtection());
            Assert.IsFalse(documentWithoutEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithoutEnforcement.Close();

            XWPFDocument documentWithReadonlyEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_readonly_no_password.docx");
            Assert.IsTrue(documentWithReadonlyEnforcement.IsEnforcedReadonlyProtection());
            Assert.IsFalse(documentWithReadonlyEnforcement.IsEnforcedFillingFormsProtection());
            Assert.IsFalse(documentWithReadonlyEnforcement.IsEnforcedCommentsProtection());
            Assert.IsFalse(documentWithReadonlyEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithReadonlyEnforcement.Close();

            XWPFDocument documentWithFillingFormsEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_forms_no_password.docx");
            Assert.IsTrue(documentWithFillingFormsEnforcement.IsEnforcedFillingFormsProtection());
            Assert.IsFalse(documentWithFillingFormsEnforcement.IsEnforcedReadonlyProtection());
            Assert.IsFalse(documentWithFillingFormsEnforcement.IsEnforcedCommentsProtection());
            Assert.IsFalse(documentWithFillingFormsEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithFillingFormsEnforcement.Close();

            XWPFDocument documentWithCommentsEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_comments_no_password.docx");
            Assert.IsFalse(documentWithCommentsEnforcement.IsEnforcedFillingFormsProtection());
            Assert.IsFalse(documentWithCommentsEnforcement.IsEnforcedReadonlyProtection());
            Assert.IsTrue(documentWithCommentsEnforcement.IsEnforcedCommentsProtection());
            Assert.IsFalse(documentWithCommentsEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithCommentsEnforcement.Close();

            XWPFDocument documentWithTrackedChangesEnforcement = XWPFTestDataSamples.OpenSampleDocument("documentProtection_trackedChanges_no_password.docx");
            Assert.IsFalse(documentWithTrackedChangesEnforcement.IsEnforcedFillingFormsProtection());
            Assert.IsFalse(documentWithTrackedChangesEnforcement.IsEnforcedReadonlyProtection());
            Assert.IsFalse(documentWithTrackedChangesEnforcement.IsEnforcedCommentsProtection());
            Assert.IsTrue(documentWithTrackedChangesEnforcement.IsEnforcedTrackedChangesProtection());
            documentWithTrackedChangesEnforcement.Close();
        }

        [Test]
        public void TestShouldEnforceForReadOnly()
        {
            //		XWPFDocument document = CreateDocumentFromSampleFile("test-data/document/documentProtection_no_protection.docx");
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            Assert.IsFalse(document.IsEnforcedReadonlyProtection());

            document.EnforceReadonlyProtection();

            Assert.IsTrue(document.IsEnforcedReadonlyProtection());
            document.Close();
        }

        [Test]
        public void TestShouldEnforceForFillingForms()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            Assert.IsFalse(document.IsEnforcedFillingFormsProtection());

            document.EnforceFillingFormsProtection();

            Assert.IsTrue(document.IsEnforcedFillingFormsProtection());
            document.Close();
        }

        [Test]
        public void TestShouldEnforceForComments()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            Assert.IsFalse(document.IsEnforcedCommentsProtection());

            document.EnforceCommentsProtection();

            Assert.IsTrue(document.IsEnforcedCommentsProtection());
            document.Close();
        }

        [Test]
        public void TestShouldEnforceForTrackedChanges()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_no_protection.docx");
            Assert.IsFalse(document.IsEnforcedTrackedChangesProtection());

            document.EnforceTrackedChangesProtection();

            Assert.IsTrue(document.IsEnforcedTrackedChangesProtection());
            document.Close();
        }

        [Test]
        public void TestShouldUnsetEnforcement()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("documentProtection_readonly_no_password.docx");
            Assert.IsTrue(document.IsEnforcedReadonlyProtection());

            document.RemoveProtectionEnforcement();

            Assert.IsFalse(document.IsEnforcedReadonlyProtection());
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

            Assert.IsTrue(doc2.IsEnforcedCommentsProtection());
            doc2.Close();
            doc1.Close();
        }

        [Test]
        public void TestUpdateFields()
        {
            XWPFDocument doc = new XWPFDocument();
            Assert.IsFalse(doc.IsEnforcedUpdateFields());
            doc.EnforceUpdateFields();
            Assert.IsTrue(doc.IsEnforcedUpdateFields());
            doc.Close();
        }

        [Test]
        public void Bug56076_read()
        {
            // test legacy xored-hashed password
            Assert.AreEqual("64CEED7E", CryptoFunctions.XorHashPassword("Example"));
            // check leading 0
            Assert.AreEqual("0005CB00", CryptoFunctions.XorHashPassword("34579"));

            // test document write protection with password
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("bug56076.docx");
            bool isValid = document.ValidateProtectionPassword("Example");
            Assert.IsTrue(isValid);
        }

        [Test]
        public void Bug56076_write()
        {
            // test document write protection with password
            XWPFDocument document = new XWPFDocument();
            document.EnforceCommentsProtection("Example", HashAlgorithm.sha512);
            document = XWPFTestDataSamples.WriteOutAndReadBack(document);
            bool isValid = document.ValidateProtectionPassword("Example");
            Assert.IsTrue(isValid);
        }
    }
}

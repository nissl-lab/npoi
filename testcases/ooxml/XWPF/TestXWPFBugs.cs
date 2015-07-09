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
namespace NPOI.XWPF
{
    using System;




    using NUnit.Framework;

    using NPOI.HSSF.Record.Crypto;
    using NPOI.POIFS.FileSystem;
    using NPOI.XWPF.UserModel;
    using TestCases;
    using System.IO;

    [TestFixture]
    public class TestXWPFBugs
    {
        /**
         * A word document that's encrypted with non-standard
         *  Encryption options, and no cspname section. See bug 53475
         */
        [Ignore("encryption function need re port from poi")]
        [Test]
        public void Test53475()
        {
            try
            {
                Biff8EncryptionKey.CurrentUserPassword = (/*setter*/"solrcell");
                FileStream file = POIDataSamples.GetDocumentInstance().GetFile("bug53475-password-is-solrcell.docx");
                NPOIFSFileSystem filesystem = new NPOIFSFileSystem(file,null, true, true);
/*
                // Check the encryption details
                EncryptionInfo info = new EncryptionInfo(filesystem);
                Assert.AreEqual(128, info.Header.KeySize);
                Assert.AreEqual(EncryptionHeader.ALGORITHM_AES_128, info.Header.Algorithm);
                Assert.AreEqual(EncryptionHeader.HASH_SHA1, info.Header.HashAlgorithm);

                // Check it can be decoded
                Decryptor d = Decryptor.GetInstance(info);
                Assert.IsTrue("Unable to Process: document is encrypted", d.VerifyPassword("solrcell"));

                // Check we can read the word document in that
                InputStream dataStream = d.GetDataStream(filesystem);
                OPCPackage opc = OPCPackage.Open(dataStream);
                XWPFDocument doc = new XWPFDocument(opc);
                XWPFWordExtractor ex = new XWPFWordExtractor(doc);
                String text = ex.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("This is password protected Word document.", text.Trim());
                ex.Close();
 */
            }
            finally
            {
                Biff8EncryptionKey.CurrentUserPassword = (/*setter*/null);
            }
        }
    }
}


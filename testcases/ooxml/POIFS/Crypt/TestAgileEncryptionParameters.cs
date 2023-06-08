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
namespace TestCases.POIFS.Crypt
{
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using TestCases;

    [TestFixture]
    public class TestAgileEncryptionParameters
    {

        static byte[] testData;

        public CipherAlgorithm ca;
        public HashAlgorithm ha;
        public ChainingMode cm;

    public static List<Object[]> data() {
            CipherAlgorithm[] caList = { CipherAlgorithm.aes128, CipherAlgorithm.aes192, CipherAlgorithm.aes256, CipherAlgorithm.rc2, CipherAlgorithm.des, CipherAlgorithm.des3 };
            HashAlgorithm[] haList = { HashAlgorithm.sha1, HashAlgorithm.sha256, HashAlgorithm.sha384, HashAlgorithm.sha512, HashAlgorithm.md5 };
            ChainingMode[] cmList = { ChainingMode.cbc, ChainingMode.cfb };

            List<Object[]> data = new List<Object[]>();
            foreach (CipherAlgorithm ca in caList) {
                foreach (HashAlgorithm ha in haList) {
                    foreach (ChainingMode cm in cmList) {
                        data.Add(new Object[] { ca, ha, cm });
                    }
                }
            }

            return data;
        }

        [OneTimeSetUp]
        public void TestData() {
            Stream testFile = POIDataSamples.GetDocumentInstance().OpenResourceAsStream("SampleDoc.docx");
            testData = IOUtils.ToByteArray(testFile);
            testFile.Close();
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestAgileEncryptionModes() {
            int maxKeyLen = Cipher.GetMaxAllowedKeyLength(ca.jceId);
            Assume.That(maxKeyLen >= ca.defaultKeySize, "Please install JCE Unlimited Strength Jurisdiction Policy files");

            MemoryStream bos = new MemoryStream();

            POIFSFileSystem fsEnc = new POIFSFileSystem();
            EncryptionInfo infoEnc = new EncryptionInfo(EncryptionMode.Agile, ca, ha, -1, -1, cm);
            Encryptor enc = infoEnc.Encryptor;
            enc.ConfirmPassword("foobaa");
            Stream os = enc.GetDataStream(fsEnc);
            os.Write(testData, 0, testData.Length);
            os.Close();
            //bos.Reset();
            bos.Seek(0, SeekOrigin.Begin);
            fsEnc.WriteFileSystem(bos);

            POIFSFileSystem fsDec = new POIFSFileSystem(new MemoryStream(bos.ToArray()));
            EncryptionInfo infoDec = new EncryptionInfo(fsDec);
            Decryptor dec = infoDec.Decryptor;
            bool passed = dec.VerifyPassword("foobaa");
            Assert.IsTrue(passed);
            Stream is1 = dec.GetDataStream(fsDec);
            byte[] actualData = IOUtils.ToByteArray(is1);
            is1.Close();
            //assertThat("Failed roundtrip - " + ca + "-" + ha + "-" + cm, testData, EqualTo(actualData));
            Assert.That(testData, Is.EqualTo(actualData), "Failed roundtrip - " + ca + "-" + ha + "-" + cm);
        }
    }
}

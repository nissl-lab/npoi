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
    using NUnit.Framework;
    using TestCases;

    [TestFixture]
    public class TestEncryptionInfo
    {

        [Test]
        public void TestEncryptionInfo1()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Assert.AreEqual(3, info.VersionMajor);
            Assert.AreEqual(2, info.VersionMinor);

            Assert.AreEqual(CipherAlgorithm.aes128, info.Header.CipherAlgorithm);
            Assert.AreEqual(HashAlgorithm.sha1, info.Header.HashAlgorithm);
            Assert.AreEqual(128, info.Header.KeySize);
            Assert.AreEqual(32, info.Verifier.EncryptedVerifierHash.Length);
            Assert.AreEqual(CipherProvider.aes, info.Header.CipherProvider);
            Assert.AreEqual("Microsoft Enhanced RSA and AES Cryptographic Provider", info.Header.CspName);

            fs.Close();
        }

        [Test]
        public void TestEncryptionInfoSHA512()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_sha512.xlsx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Assert.AreEqual(4, info.VersionMajor);
            Assert.AreEqual(4, info.VersionMinor);

            Assert.AreEqual(CipherAlgorithm.aes256, info.Header.CipherAlgorithm);
            Assert.AreEqual(HashAlgorithm.sha512, info.Header.HashAlgorithm);
            Assert.AreEqual(256, info.Header.KeySize);
            Assert.AreEqual(64, info.Verifier.EncryptedVerifierHash.Length);
            Assert.AreEqual(CipherProvider.aes, info.Header.CipherProvider);
            //        Assert.AreEqual("Microsoft Enhanced RSA and AES Cryptographic Provider", info.Header.CspName);

            fs.Close();
        }
    }
}
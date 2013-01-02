/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Crypt;

namespace TestCases.POIFS.Crypt
{
    /**
     *  @author Maxim Valyanskiy
     */
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

            Assert.AreEqual(EncryptionHeader.ALGORITHM_AES_128, info.Header.Algorithm);
            Assert.AreEqual(EncryptionHeader.HASH_SHA1, info.Header.HashAlgorithm);
            Assert.AreEqual(128, info.Header.KeySize);
            Assert.AreEqual(EncryptionHeader.PROVIDER_AES, info.Header.ProviderType);
            Assert.AreEqual("Microsoft Enhanced RSA and AES Cryptographic Provider", info.Header.CspName);

            Assert.AreEqual(32, info.Verifier.VerifierHash.Length);
        }
    }
}

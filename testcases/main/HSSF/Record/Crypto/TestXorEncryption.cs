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

namespace TestCases.HSSF.Record.Crypto
{
    using System;
    using NPOI.HSSF.Record.Crypto;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NUnit.Framework;
    using TestCases.HSSF;

    [TestFixture]
    public class TestXorEncryption
    {

        private static HSSFTestDataSamples samples = new HSSFTestDataSamples();

        // to not affect other tests Running in the same JVM
        [TearDown]
        public void ResetPassword() {
            Biff8EncryptionKey.CurrentUserPassword = (/*setter*/null);
        }

        [Test]
        public void TestXorEncryption1() {
            // Xor-Password: abc
            // 2.5.343 XORObfuscation
            // key = 20810
            // verifier = 52250
            int verifier = CryptoFunctions.CreateXorVerifier1("abc");
            int key = CryptoFunctions.CreateXorKey1("abc");
            Assert.AreEqual(20810, key);
            Assert.AreEqual(52250, verifier);

            byte[] xorArrAct = CryptoFunctions.CreateXorArray1("abc");
            byte[] xorArrExp = HexRead.ReadFromString("AC-CC-A4-AB-D6-BA-C3-BA-D6-A3-2B-45-D3-79-29-BB");

            //assertThat(xorArrExp, EqualTo(xorArrAct));
            Assert.That(xorArrExp, Is.EqualTo(xorArrAct));
        }

        [Test]
        [Ignore("not implemented")]
        public void TestUserFile() {
            Biff8EncryptionKey.CurrentUserPassword = (/*setter*/"abc");
            NPOIFSFileSystem fs = new NPOIFSFileSystem(HSSFTestDataSamples.GetSampleFile("xor-encryption-abc.xls"), true);
            IWorkbook hwb = new HSSFWorkbook(fs.Root, true);

            ISheet sh = hwb.GetSheetAt(0);
            Assert.AreEqual(1.0, sh.GetRow(0).GetCell(0).NumericCellValue, 0.0);
            Assert.AreEqual(2.0, sh.GetRow(1).GetCell(0).NumericCellValue, 0.0);
            Assert.AreEqual(3.0, sh.GetRow(2).GetCell(0).NumericCellValue, 0.0);

            fs.Close();
        }
    }

}
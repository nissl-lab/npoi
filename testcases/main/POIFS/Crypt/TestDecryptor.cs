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
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Crypt;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;
using NUnit.Framework;
using NPOI.Util;
namespace TestCases.POIFS.Crypt
{
    /**
 *  @author Maxim Valyanskiy
 *  @author Gary King
 */
    [TestFixture]
    public class TestDecryptor
    {
        [Test]
        public void TestPasswordVerification()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            Assert.IsTrue(d.VerifyPassword(Decryptor.DEFAULT_PASSWORD));
        }
        [Test]
        public void TestDecrypt()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            d.VerifyPassword(Decryptor.DEFAULT_PASSWORD);

            ZipOk(fs, d);
        }
        [Test]
        public void TestAgile()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_agile.docx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Assert.IsTrue(info.VersionMajor == 4 && info.VersionMinor == 4);

            Decryptor d = Decryptor.GetInstance(info);

            Assert.IsTrue(d.VerifyPassword(Decryptor.DEFAULT_PASSWORD));

            ZipOk(fs, d);
        }

        private void ZipOk(POIFSFileSystem fs, Decryptor d)
        {
            ZipInputStream zin = new ZipInputStream(d.GetDataStream(fs));
            while (true)
            {
                ZipEntry entry = zin.GetNextEntry();
                if (entry == null)
                {
                    break;
                }

                while (zin.Available > 0)
                {
                    zin.Skip(zin.Available);
                }
            }
        }
        [Test]
        public void TestDataLength()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_agile.docx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            d.VerifyPassword(Decryptor.DEFAULT_PASSWORD);

            Stream is1 = d.GetDataStream(fs);

            long len = d.Length;
            Assert.AreEqual(12810, len);

            byte[] buf = new byte[(int)len];

            is1.Read(buf, 0, buf.Length);

            ZipInputStream zin = new ZipInputStream(new ByteArrayInputStream(buf));

            while (true)
            {
                ZipEntry entry = zin.GetNextEntry();
                if (entry == null)
                {
                    break;
                }

                while (zin.Available > 0)
                {
                    zin.Skip(zin.Available);
                }
            }
        }
    }
}

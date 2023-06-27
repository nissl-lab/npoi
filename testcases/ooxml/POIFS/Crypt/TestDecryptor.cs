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
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NPOI.XSSF;
    using NUnit.Framework;
    using System.IO;
    using TestCases;

    /**
     *  @author Maxim Valyanskiy
     *  @author Gary King
     */
    [TestFixture]
    public class TestDecryptor
    {
        [Test]
        public void PasswordVerification()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            Assert.IsTrue(d.VerifyPassword(Decryptor.DEFAULT_PASSWORD));
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void Decrypt()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            d.VerifyPassword(Decryptor.DEFAULT_PASSWORD);

            ZipOk(fs.Root, d);
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void Agile()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_agile.docx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Assert.IsTrue(info.VersionMajor == 4 && info.VersionMinor == 4);

            Decryptor d = Decryptor.GetInstance(info);

            Assert.IsTrue(d.VerifyPassword(Decryptor.DEFAULT_PASSWORD));

            ZipOk(fs.Root, d);
        }

        private void ZipOk(DirectoryNode root, Decryptor d)
        {
            ZipInputStream zin = new ZipInputStream(d.GetDataStream(root));

            while (true)
            {
                ZipEntry entry = zin.GetNextEntry();
                if (entry == null) break;
                // crc32 is Checked within zip-stream
                if (entry.IsDirectory) continue;
                zin.Skip(entry.Size);
                byte[] buf = new byte[10];
                int ReadBytes = zin.Read(buf, 0, buf.Length);
                // zin.Available() doesn't work for entries
                Assert.AreEqual(-1, ReadBytes, "size failed for " + entry.Name);
            }

            zin.Close();
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void DataLength()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_agile.docx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            d.VerifyPassword(Decryptor.DEFAULT_PASSWORD);

            Stream is1 = d.GetDataStream(fs);

            long len = d.GetLength();
            Assert.AreEqual(12810, len);

            byte[] buf = new byte[(int)len];

            is1.Read(buf, 0, buf.Length);

            ZipInputStream zin = new ZipInputStream(new MemoryStream(buf));

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
        public void Bug57080()
        {
            // the test file Contains a wrong ole entry size, produced by extenxls
            // the fix limits the available size and tries to read all entries 
            FileStream f = POIDataSamples.GetPOIFSInstance().GetFile("extenxls_pwd123.xlsx");
            NPOIFSFileSystem fs = new NPOIFSFileSystem(f, true);
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor d = Decryptor.GetInstance(info);
            d.VerifyPassword("pwd123");
            MemoryStream bos = new MemoryStream();
            ZipInputStream zis = new ZipInputStream(d.GetDataStream(fs));
            ZipEntry ze;
            while ((ze = zis.GetNextEntry()) != null)
            {
                //bos.Reset();
                bos.Seek(0, SeekOrigin.Begin);
                bos.SetLength(0);
                IOUtils.Copy(zis, bos);
                Assert.AreEqual(ze.Size, bos.Length);
            }

            zis.Close();
            fs.Close();
        }

        [Test]
        public void Test58616()
        {
            POIFSFileSystem pfs = new POIFSFileSystem(XSSFTestDataSamples.GetSampleFile("58616.xlsx"));
            EncryptionInfo info = new EncryptionInfo(pfs);
            Decryptor dec = Decryptor.GetInstance(info);
            //dec.VerifyPassword(null);
            dec.GetDataStream(pfs);
        }
    }
}
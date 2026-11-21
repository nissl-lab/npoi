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
    using NUnit.Framework.Legacy;
    using System.Diagnostics;
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

            ClassicAssert.IsTrue(d.VerifyPassword(Decryptor.DEFAULT_PASSWORD));
        }

        [Test]
        public void Decrypt()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            d.VerifyPassword(Decryptor.DEFAULT_PASSWORD);

            using(ZipInputStream zin = new ZipInputStream(d.GetDataStream(fs.Root)))
            {
                ZipOk(zin);
            }
        }

        [Test]
        public void Agile()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_agile.docx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            ClassicAssert.IsTrue(info.VersionMajor == 4 && info.VersionMinor == 4);

            Decryptor d = Decryptor.GetInstance(info);

            ClassicAssert.IsTrue(d.VerifyPassword(Decryptor.DEFAULT_PASSWORD));

            using(ZipInputStream zin = new ZipInputStream(d.GetDataStream(fs.Root)))
            {
                ZipOk(zin);
            }
        }

        private void ZipOk(ZipInputStream zin)
        {
            // Fully read each ZIP entry produced by the decrypted stream.
            // Old logic used Skip(entry.Size) and then expected an extra Read() == -1,
            // which could under/over-read and fail when entry.Size == -1 (unknown).
            byte[] buffer = new byte[4096];
            ZipEntry entry;
            while((entry = zin.GetNextEntry()) != null)
            {
                Debug.WriteLine(entry.Name);
                if(entry.IsDirectory)
                    continue;

                long total = 0;
                int read;
                // Read until end-of-entry (SharpZipLib returns 0 when entry finished)
                while((read = zin.Read(buffer, 0, buffer.Length)) > 0)
                {
                    total += read;
                }

                // If size is known (>=0) assert it matches uncompressed bytes read
                if(entry.Size >= 0)
                {
                    ClassicAssert.AreEqual(entry.Size, total, "Size mismatch for " + entry.Name);
                }
            }
        }

        [Test]
        public void DataLength()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_agile.docx"));

            EncryptionInfo info = new EncryptionInfo(fs);

            Decryptor d = Decryptor.GetInstance(info);

            d.VerifyPassword(Decryptor.DEFAULT_PASSWORD);

            Stream is1 = d.GetDataStream(fs);

            long len = d.GetLength();
            ClassicAssert.AreEqual(12810, len);

            byte[] buf = new byte[(int)len];

            is1.Read(buf, 0, buf.Length);

            ZipInputStream zin = new ZipInputStream(new MemoryStream(buf));
            ZipOk(zin);
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
            while((ze = zis.GetNextEntry()) != null)
            {
                //bos.Reset();
                bos.Seek(0, SeekOrigin.Begin);
                bos.SetLength(0);
                IOUtils.Copy(zis, bos);
                ClassicAssert.AreEqual(ze.Size, bos.Length);
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
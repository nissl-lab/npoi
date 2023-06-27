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
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXml4Net.Util;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NPOI.XSSF;
    using NUnit.Framework;
    using Org.BouncyCastle.Security;
    using System;
    using System.Collections;
    using System.IO;

    [TestFixture]
    public class TestSecureTempZip
    {
        /**
         * Test case for #59841 - this is an example on how to use encrypted temp files,
         * which are streamed into POI opposed to having everything in memory
         */
        [Ignore("not implemented")]
        public void ProtectedTempZip()
        {
            FileInfo tmpFile = TempFile.CreateTempFile("protectedXlsx", ".zip");
            FileInfo tikaProt = XSSFTestDataSamples.GetSampleFile("protected_passtika.xlsx");
            FileInputStream fis = new FileInputStream(tikaProt.Open(FileMode.Open));
            POIFSFileSystem poifs = new POIFSFileSystem(fis);
            EncryptionInfo ei = new EncryptionInfo(poifs);
            Decryptor dec = ei.Decryptor;
            bool passOk = dec.VerifyPassword("tika");
            Assert.IsTrue(passOk);

            // generate session key
            SecureRandom sr = new SecureRandom();
            byte[] ivBytes = new byte[16], keyBytes = new byte[16];
            sr.NextBytes(ivBytes);
            sr.NextBytes(keyBytes);

            // extract encrypted ooxml file and write to custom encrypted zip file 
            InputStream is1 = dec.GetDataStream(poifs);
            CopyToFile(is1, tmpFile, CipherAlgorithm.aes128, keyBytes, ivBytes);
            is1.Close();

            // provide ZipEntrySource to poi which decrypts on the fly
            ZipEntrySource source = fileToSource(tmpFile, CipherAlgorithm.aes128, keyBytes, ivBytes);

            // test the source
            OPCPackage opc = OPCPackage.Open(source);
            // String expected = "This is an Encrypted Excel spreadsheet.";

            //XSSFEventBasedExcelExtractor extractor = new XSSFEventBasedExcelExtractor(opc);
            //extractor.IncludeSheetNames = (/*setter*/false);
            //String txt = extractor.Text;
            //Assert.AreEqual(expected, txt.Trim());

            //XSSFWorkbook wb = new XSSFWorkbook(opc);
            //txt = wb.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue;
            //Assert.AreEqual(expected, txt);

            //extractor.Close();

            //wb.Close();
            opc.Close();
            source.Close();
            poifs.Close();
            fis.Close();
            tmpFile.Delete();

            throw new NotImplementedException();
        }

        private void CopyToFile(InputStream is1, FileInfo tmpFile, CipherAlgorithm cipherAlgorithm, byte[] keyBytes, byte[] ivBytes)
        {
            SecretKeySpec skeySpec = new SecretKeySpec(keyBytes, cipherAlgorithm.jceId);
            Cipher ciEnc = CryptoFunctions.GetCipher(skeySpec, cipherAlgorithm, ChainingMode.cbc, ivBytes, Cipher.ENCRYPT_MODE, "PKCS5PAdding");

            ZipInputStream zis = new ZipInputStream(is1);
            //FileOutputStream fos = new FileOutputStream(tmpFile);
            //ZipOutputStream zos = new ZipOutputStream(fos);

            //ZipEntry ze;
            //while ((ze = zis.NextEntry) != null)
            //{
            //    // the cipher output stream pads the data, therefore we can't reuse the ZipEntry with Set sizes
            //    // as those will be validated upon close()
            //    ZipEntry zeNew = new ZipEntry(ze.Name);
            //    zeNew.Comment = (/*setter*/ze.Comment);
            //    zeNew.Extra = (/*setter*/ze.Extra);
            //    zeNew.Time = (/*setter*/ze.Time);
            //    // zeNew.Method=(/*setter*/ze.Method);
            //    zos.PutNextEntry(zeNew);
            //    FilterOutputStream fos2 = new FilterOutputStream(zos)
            //    {
            //        // don't close underlying ZipOutputStream
            //        public void close() { }
            //};
            //CipherOutputStream cos = new CipherOutputStream(fos2, ciEnc);
            //    IOUtils.Copy(zis, cos);
            //    cos.Close();
            //    fos2.Close();
            //    zos.CloseEntry();
            //    zis.CloseEntry();
            //}
            //zos.Close();
            //fos.Close();
            //zis.Close();
            throw new NotImplementedException();
        }

        private ZipEntrySource fileToSource(FileInfo tmpFile, CipherAlgorithm cipherAlgorithm, byte[] keyBytes, byte[] ivBytes)
        {
            SecretKeySpec skeySpec = new SecretKeySpec(keyBytes, cipherAlgorithm.jceId);
            Cipher ciDec = CryptoFunctions.GetCipher(skeySpec, cipherAlgorithm, ChainingMode.cbc, ivBytes, Cipher.DECRYPT_MODE, "PKCS5PAdding");
            ZipFile zf = new ZipFile(tmpFile.FullName);
            return new AesZipFileZipEntrySource(zf, ciDec);
        }

        class AesZipFileZipEntrySource : ZipEntrySource
        {
            ZipFile zipFile;
            Cipher ci;
            bool closed;

            public AesZipFileZipEntrySource(ZipFile zipFile, Cipher ci)
            {
                this.zipFile = zipFile;
                this.ci = ci;
                this.closed = false;
            }

            /**
             * Note: the file sizes are rounded up to the next cipher block size,
             * so don't rely on file sizes of these custom encrypted zip file entries!
             */
            public IEnumerator Entries
            {
                get
                {
                    return zipFile.GetEnumerator();
                }

            }

            public Stream GetInputStream(ZipEntry entry)
            {
                Stream is1 = zipFile.GetInputStream(entry);
                throw new NotImplementedException();
                //return new CipherInputStream(is1, ci);
            }


            public void Close()
            {
                zipFile.Close();
                closed = true;
            }


            public bool IsClosed
            {
                get
                {
                    return closed;
                }
            }
        }
    }

}
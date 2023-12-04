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
    using System;
    using System.Collections.Generic;
    using System.IO;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.Agile;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using TestCases;

    [TestFixture]
    public class TestEncryptor
    {
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void BinaryRC4Encryption()
        {
            // please contribute a real sample file, which is binary rc4 encrypted
            // ... at least the output can be opened in Excel Viewer 
            String password = "pass";

            Stream is1 = POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("SimpleMultiCell.xlsx");
            MemoryStream payloadExpected = new MemoryStream();
            IOUtils.Copy(is1, payloadExpected);
            is1.Close();

            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo ei = new EncryptionInfo(EncryptionMode.BinaryRC4);
            Encryptor enc = ei.Encryptor;
            enc.ConfirmPassword(password);

            Stream os = enc.GetDataStream(fs.Root);
            payloadExpected.WriteTo(os);
            os.Close();

            MemoryStream bos = new MemoryStream();
            fs.WriteFileSystem(bos);

            fs = new POIFSFileSystem(new MemoryStream(bos.ToArray()));
            ei = new EncryptionInfo(fs);
            Decryptor dec = ei.Decryptor;
            bool b = dec.VerifyPassword(password);
            Assert.IsTrue(b);

            MemoryStream payloadActual = new MemoryStream();
            is1 = dec.GetDataStream(fs.Root);
            IOUtils.Copy(is1, payloadActual);
            is1.Close();

            Assert.IsTrue(Arrays.Equals(payloadExpected.ToArray(), payloadActual.ToArray()));
            //assertArrayEquals(payloadExpected.ToArray(), payloadActual.ToArray());
        }

        [Test]
        public void AgileEncryption()
        {
            int maxKeyLen = Cipher.GetMaxAllowedKeyLength("AES");
            Assume.That(maxKeyLen == 2147483647, "Please install JCE Unlimited Strength Jurisdiction Policy files for AES 256");

            FileStream file = POIDataSamples.GetDocumentInstance().GetFile("bug53475-password-is-pass.docx");
            String pass = "pass";
            NPOIFSFileSystem nfs = new NPOIFSFileSystem(file);

            // Check the encryption details
            EncryptionInfo infoExpected = new EncryptionInfo(nfs);
            Decryptor decExpected = Decryptor.GetInstance(infoExpected);
            bool passed = decExpected.VerifyPassword(pass);
            Assert.IsTrue(passed, "Unable to Process: document is encrypted");

            // extract the payload
            Stream is1 = decExpected.GetDataStream(nfs);
            byte[] payloadExpected = IOUtils.ToByteArray(is1);
            is1.Close();

            long decPackLenExpected = decExpected.GetLength();
            Assert.AreEqual(decPackLenExpected, payloadExpected.Length);

            is1 = nfs.Root.CreateDocumentInputStream(Decryptor.DEFAULT_POIFS_ENTRY);
            ///is1 = new BoundedInputStream(is1, is1.Available() - 16); // ignore pAdding block
            ///throw new NotImplementedException(BoundedInputStream); 
            byte[] encPackExpected = IOUtils.ToByteArray(is1);
            is1.Close();

            // listDir(nfs.Root, "orig", "");

            nfs.Close();

            // check that same verifier/salt lead to same hashes
            byte[] verifierSaltExpected = infoExpected.Verifier.Salt;
            byte[] verifierExpected = decExpected.GetVerifier();
            byte[] keySalt = infoExpected.Header.KeySalt;
            byte[] keySpec = decExpected.GetSecretKey().GetEncoded();
            byte[] integritySalt = decExpected.GetIntegrityHmacKey();
            // the hmacs of the file always differ, as we use PKCS5-pAdding to pad the bytes
            // whereas office just uses random bytes
            // byte integrityHash[] = d.IntegrityHmacValue;

            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo infoActual = new EncryptionInfo(
                  EncryptionMode.Agile
                , infoExpected.Verifier.CipherAlgorithm
                , infoExpected.Verifier.HashAlgorithm
                , infoExpected.Header.KeySize
                , infoExpected.Header.BlockSize
                , infoExpected.Verifier.ChainingMode
            );

            Encryptor e = Encryptor.GetInstance(infoActual);
            e.ConfirmPassword(pass, keySpec, keySalt, verifierExpected, verifierSaltExpected, integritySalt);

            Stream os = e.GetDataStream(fs);
            IOUtils.Copy(new MemoryStream(payloadExpected), os);
            os.Close();

            MemoryStream bos = new MemoryStream();
            fs.WriteFileSystem(bos);

            nfs = new NPOIFSFileSystem(new MemoryStream(bos.ToArray()));
            infoActual = new EncryptionInfo(nfs.Root);
            Decryptor decActual = Decryptor.GetInstance(infoActual);
            passed = decActual.VerifyPassword(pass);
            Assert.IsTrue(passed, "Unable to Process: document is encrypted");

            // extract the payload
            is1 = decActual.GetDataStream(nfs);
            byte[] payloadActual = IOUtils.ToByteArray(is1);
            is1.Close();

            long decPackLenActual = decActual.GetLength();

            is1 = nfs.Root.CreateDocumentInputStream(Decryptor.DEFAULT_POIFS_ENTRY);
            ///is1 = new BoundedInputStream(is1, is1.Available() - 16); // ignore pAdding block
            ///throw new NotImplementedException(BoundedInputStream);
            byte[] encPackActual = IOUtils.ToByteArray(is1);
            is1.Close();

            // listDir(nfs.Root, "copy", "");

            nfs.Close();

            AgileEncryptionHeader aehExpected = (AgileEncryptionHeader)infoExpected.Header;
            AgileEncryptionHeader aehActual = (AgileEncryptionHeader)infoActual.Header;
            CollectionAssert.AreEqual(aehExpected.GetEncryptedHmacKey(), aehActual.GetEncryptedHmacKey());
            Assert.AreEqual(decPackLenExpected, decPackLenActual);
            CollectionAssert.AreEqual(payloadExpected, payloadActual);
            CollectionAssert.AreEqual(encPackExpected, encPackActual);
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void StandardEncryption()
        {
            FileStream file = POIDataSamples.GetDocumentInstance().GetFile("bug53475-password-is-solrcell.docx");
            String pass = "solrcell";

            NPOIFSFileSystem nfs = new NPOIFSFileSystem(file);

            // Check the encryption details
            EncryptionInfo infoExpected = new EncryptionInfo(nfs);
            Decryptor d = Decryptor.GetInstance(infoExpected);
            bool passed = d.VerifyPassword(pass);
            Assert.IsTrue(passed, "Unable to Process: document is encrypted");

            // extract the payload
            MemoryStream bos = new MemoryStream();
            Stream is1 = d.GetDataStream(nfs);
            IOUtils.Copy(is1, bos);
            is1.Close();
            nfs.Close();
            byte[] payloadExpected = bos.ToArray();

            // check that same verifier/salt lead to same hashes
            byte[] verifierSaltExpected = infoExpected.Verifier.Salt;
            byte[] verifierExpected = d.GetVerifier();
            byte[] keySpec = d.GetSecretKey().GetEncoded();
            byte[] keySalt = infoExpected.Header.KeySalt;


            EncryptionInfo infoActual = new EncryptionInfo(
                  EncryptionMode.Standard
                , infoExpected.Verifier.CipherAlgorithm
                , infoExpected.Verifier.HashAlgorithm
                , infoExpected.Header.KeySize
                , infoExpected.Header.BlockSize
                , infoExpected.Verifier.ChainingMode
            );

            Encryptor e = Encryptor.GetInstance(infoActual);
            e.ConfirmPassword(pass, keySpec, keySalt, verifierExpected, verifierSaltExpected, null);

            CollectionAssert.AreEqual(infoExpected.Verifier.EncryptedVerifier, infoActual.Verifier.EncryptedVerifier);
            CollectionAssert.AreEqual(infoExpected.Verifier.EncryptedVerifierHash, infoActual.Verifier.EncryptedVerifierHash);

            // now we use a newly generated salt/verifier and check
            // if the file content is still the same 

            infoActual = new EncryptionInfo(
                  EncryptionMode.Standard
                , infoExpected.Verifier.CipherAlgorithm
                , infoExpected.Verifier.HashAlgorithm
                , infoExpected.Header.KeySize
                , infoExpected.Header.BlockSize
                , infoExpected.Verifier.ChainingMode
            );

            e = Encryptor.GetInstance(infoActual);
            e.ConfirmPassword(pass);

            POIFSFileSystem fs = new POIFSFileSystem();
            Stream os = e.GetDataStream(fs);
            IOUtils.Copy(new MemoryStream(payloadExpected), os);
            os.Close();

            bos.Seek(0, SeekOrigin.Begin); //bos.Reset();
            fs.WriteFileSystem(bos);

            ByteArrayInputStream bis = new ByteArrayInputStream(bos.ToArray());

            // FileOutputStream fos = new FileOutputStream("encrypted.docx");
            // IOUtils.Copy(bis, fos);
            // fos.Close();
            // bis.Reset();

            nfs = new NPOIFSFileSystem(bis);
            infoExpected = new EncryptionInfo(nfs);
            d = Decryptor.GetInstance(infoExpected);
            passed = d.VerifyPassword(pass);
            Assert.IsTrue(passed, "Unable to Process: document is encrypted");

            bos.Seek(0, SeekOrigin.Begin); //bos.Reset();
            is1 = d.GetDataStream(nfs);
            IOUtils.Copy(is1, bos);
            is1.Close();
            nfs.Close();
            byte[] payloadActual = bos.ToArray();

            CollectionAssert.AreEqual(payloadExpected, payloadActual);
            //assertArrayEquals(payloadExpected, payloadActual);
        }

        /**
         * Ensure we can encrypt a package that is missing the Core
         *  Properties, eg one from dodgy versions of Jasper Reports 
         * See https://github.com/nestoru/xlsxenc/ and
         * http://stackoverflow.com/questions/28593223
         */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void EncryptPackageWithoutCoreProperties()
        {
            // Open our file without core properties
            FileStream inp = POIDataSamples.GetOpenXML4JInstance().GetFile("OPCCompliance_NoCoreProperties.xlsx");
            OPCPackage pkg = OPCPackage.Open(inp.Name);

            // It doesn't have any core properties yet
            Assert.AreEqual(0, pkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            Assert.IsNotNull(pkg.GetPackageProperties());
            Assert.IsNotNull(pkg.GetPackageProperties().GetLanguageProperty());
            //Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty().GetValue());

            // Encrypt it
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.Agile);
            NPOIFSFileSystem fs = new NPOIFSFileSystem();

            Encryptor enc = info.Encryptor;
            enc.ConfirmPassword("password");
            OutputStream os = enc.GetDataStream(fs);
            pkg.Save(os);
            pkg.Revert();

            // Save the resulting OLE2 document, and re-open it
            MemoryStream baos = new MemoryStream();
            fs.WriteFileSystem(baos);

            MemoryStream bais = new MemoryStream(baos.ToArray());
            NPOIFSFileSystem inpFS = new NPOIFSFileSystem(bais);

            // Check we can decrypt it
            info = new EncryptionInfo(inpFS);
            Decryptor d = Decryptor.GetInstance(info);
            Assert.AreEqual(true, d.VerifyPassword("password"));

            OPCPackage inpPkg = OPCPackage.Open(d.GetDataStream(inpFS));

            // Check it now has empty core properties
            Assert.AreEqual(1, inpPkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            Assert.IsNotNull(inpPkg.GetPackageProperties());
            Assert.IsNotNull(inpPkg.GetPackageProperties().GetLanguageProperty());
            //Assert.IsNull(inpPkg.PackageProperties.LanguageProperty.Value);
        }

        [Test]
        [Ignore("poi")]
        public void InPlaceReWrite()
        {
            FileInfo f = TempFile.CreateTempFile("protected_agile", ".docx");
            // File f = new File("protected_agile.docx");
            FileStream fos = f.Create();
            Stream fis = POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protected_agile.docx");
            IOUtils.Copy(fis, fos);
            fis.Close();
            fos.Close();

            NPOIFSFileSystem fs = new NPOIFSFileSystem(f, false);

            // decrypt the protected file - in this case it was encrypted with the default password
            EncryptionInfo encInfo = new EncryptionInfo(fs);
            Decryptor d = encInfo.Decryptor;
            bool b = d.VerifyPassword(Decryptor.DEFAULT_PASSWORD);
            Assert.IsTrue(b);

            // do some strange things with it ;)
            XWPFDocument docx = new XWPFDocument(d.GetDataStream(fs));
            docx.GetParagraphArray(0).InsertNewRun(0).SetText("POI was here! All your base are belong to us!");
            docx.GetParagraphArray(0).InsertNewRun(1).AddBreak();

            // and encrypt it again
            Encryptor e = encInfo.Encryptor;
            e.ConfirmPassword("AYBABTU");
            docx.Write(e.GetDataStream(fs));

            fs.Close();
        }


        private void ListEntry(DocumentNode de, string ext, string path)
        {
            path += "\\" + de.Name.Replace("[\\p{Cntrl}]", "_");
            Console.WriteLine(ext + ": " + path + " (" + de.Size + " bytes)");

            string name = de.Name.Replace("[\\p{Cntrl}]", "_");

            Stream is1 = ((DirectoryNode)de.Parent).CreateDocumentInputStream(de);
            FileStream fos = new FileStream("solr." + name + "." + ext, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            IOUtils.Copy(is1, fos);
            fos.Close();
            is1.Close();
        }


        private void ListDir(DirectoryNode dn, string ext, string path)
        {
            path += "\\" + dn.Name.Replace('\u0006', '_');
            Console.WriteLine(ext + ": " + path + " (" + dn.StorageClsid + ")");

            IEnumerator<Entry> iter = dn.Entries;
            while (iter.MoveNext())
            {
                Entry ent = iter.Current;
                if (ent is DirectoryNode)
                {
                    ListDir((DirectoryNode)ent, ext, path);
                }
                else
                {
                    ListEntry((DocumentNode)ent, ext, path);
                }
            }
        }
    }

}
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
    using System.IO;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.Agile;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NUnit.Framework;
    using Org.BouncyCastle.X509;
    using TestCases;


    /**
     * @see <a href="http://stackoverflow.com/questions/1615871/creating-an-x509-certificate-in-java-without-bouncycastle">creating a self-signed certificate</a> 
     */
    [TestFixture]
    public class TestCertificateEncryption
    {
        /**
         * how many days from now the Certificate is valid for
         */
        // static int days = 1000;
        /**
         * the signing algorithm, eg "SHA1withRSA"
         */
        // static String algorithm = "SHA1withRSA";
        // static String password = "foobaa";
        // static String certAlias = "poitest";
        /**
         * the X.509 Distinguished Name, eg "CN=Test, L=London, C=GB"
         */
        // static String certDN = "CN=poitest";
        // static File pfxFile = TempFile.CreateTempFile("poitest", ".pfx");
        // static byte[] pfxFileBytes;

        public class CertData
        {
            public KeyPair keypair;
            public X509Certificate x509;
        }

        /** 
         * Create a self-signed X.509 Certificate
         * 
         * The keystore generation / loading is split, because normally the keystore would
         * already exist.
         */
        /* @BeforeClass
        public static void InitKeystore() {
            CertData certData = new CertData();

            KeyPairGenerator keyGen = KeyPairGenerator.GetInstance("RSA");
            keyGen.Initialize(1024);
            certData.keypair = keyGen.GenerateKeyPair();
            PrivateKey privkey = certData.keypair.Private;
            PublicKey publkey = certData.keypair.Public;

            X509CertInfo info = new X509CertInfo();
            Date from = new Date();
            Date to = new Date(from.Time + days * 86400000l);
            CertificateValidity interval = new CertificateValidity(from, to);
            Bigint sn = new Bigint(64, new SecureRandom());
            X500Name owner = new X500Name(certDN);

            info.Set(X509CertInfo.VALIDITY, interval);
            info.Set(X509CertInfo.SERIAL_NUMBER, new CertificateSerialNumber(sn));
            info.Set(X509CertInfo.SUBJECT, new CertificateSubjectName(owner));
            info.Set(X509CertInfo.ISSUER, new CertificateIssuerName(owner));
            info.Set(X509CertInfo.KEY, new CertificateX509Key(publkey));
            info.Set(X509CertInfo.VERSION, new CertificateVersion(CertificateVersion.V3));
            AlgorithmId algo = new AlgorithmId(AlgorithmId.md5WithRSAEncryption_oid);
            info.Set(X509CertInfo.ALGORITHM_ID, new CertificateAlgorithmId(algo));

            // Sign the cert to identify the algorithm that's used.
            X509CertImpl cert = new X509CertImpl(info);
            cert.Sign(privkey, algorithm);

            // Update the algorith, and resign.
            algo = (AlgorithmId)cert.Get(X509CertImpl.SIG_ALG);
            info.Set(CertificateAlgorithmId.NAME + "." + CertificateAlgorithmId.ALGORITHM, algo);
            cert = new X509CertImpl(info);
            cert.Sign(privkey, algorithm);
            certData.x509 = cert;

            KeyStore keystore = KeyStore.GetInstance("PKCS12");
            keystore.Load(null, password.ToCharArray());
            keystore.SetKeyEntry(certAlias, certData.keypair.Private, password.ToCharArray(), new Certificate[]{certData.x509});
            MemoryStream bos = new MemoryStream();
            keystore.Store(bos, password.ToCharArray());
            pfxFileBytes = bos.ToByteArray();
        } */

        public CertData loadKeystore()
        {
            //KeyStore keystore = KeyStore.GetInstance("PKCS12");

            //// InputStream fis = new MemoryStream(pfxFileBytes);
            //Stream fis = POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("poitest.pfx");
            //keystore.Load(fis, password.ToCharArray());
            //fis.Close();

            //X509Certificate x509 = (X509Certificate)keystore.GetCertificate(certAlias);
            //PrivateKey privateKey = (PrivateKey)keystore.GetKey(certAlias, password.ToCharArray());
            //PublicKey publicKey = x509.PublicKey;

            //CertData certData = new CertData();
            //certData.keypair = new KeyPair(publicKey, privateKey);
            //certData.x509 = x509;

            //return certData;
            throw new NotImplementedException();
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestCertificateEncryption1()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.Agile, CipherAlgorithm.aes128, HashAlgorithm.sha1, -1, -1, ChainingMode.cbc);
            AgileEncryptionVerifier aev = (AgileEncryptionVerifier)info.Verifier;
            CertData certData = loadKeystore();
            aev.AddCertificate(certData.x509);

            Encryptor enc = info.Encryptor;
            enc.ConfirmPassword("foobaa");

            FileStream file = POIDataSamples.GetDocumentInstance().GetFile("VariousPictures.docx");
            //InputStream fis = new FileInputStream(file);
            byte[] byteExpected = IOUtils.ToByteArray(file);
            //fis.Close();

            Stream os = enc.GetDataStream(fs);
            IOUtils.Copy(new MemoryStream(byteExpected), os);
            os.Close();

            MemoryStream bos = new MemoryStream();
            fs.WriteFileSystem(bos);
            bos.Close();

            fs = new POIFSFileSystem(new MemoryStream(bos.ToArray()));
            info = new EncryptionInfo(fs);
            AgileDecryptor agDec = (AgileDecryptor)info.Decryptor;
            bool passed = agDec.VerifyPassword(certData.keypair, certData.x509);
            Assert.IsTrue(passed, "certificate verification failed");

            Stream fis = agDec.GetDataStream(fs);
            byte[] byteActual = IOUtils.ToByteArray(fis);
            fis.Close();

            Assert.That(byteExpected, Is.EqualTo(byteActual));
        }
    }

}
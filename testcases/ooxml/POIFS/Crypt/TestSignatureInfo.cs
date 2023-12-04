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

/* ====================================================================
   This product Contains an ASLv2 licensed version of the OOXML signer
   package from the eID Applet project
   http://code.google.com/p/eid-applet/source/browse/tRunk/README.txt  
   Copyright (C) 2008-2014 FedICT.
   ================================================================= */
using System.IO;

namespace TestCases.POIFS.Crypt
{
    using NPOI.OpenXml4Net.OPC;
    using NPOI.POIFS.Crypt.Dsig;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using System.Security.Cryptography.X509Certificates;
    using TestCases;

    [TestFixture]
    public class TestSignatureInfo
    {
        private static POIDataSamples testdata = POIDataSamples.GetXmlDSignInstance();

        //private static Calendar cal;
        //private KeyPair keyPair = null;

        [SetUp]
        public static void InitBouncy()
        {
            throw new NotImplementedException();
            //CryptoFunctions.RegisterBouncyCastle();

            ///*** TODO : Set cal to now ... only Set to fixed date for debugging ... */
            //cal = Calendar.Instance;
            //cal.Clear();
            //cal.TimeZone = (/*setter*/TimeZone.GetTimeZone("UTC"));
            //cal.Set(2014, 7, 6, 21, 42, 12);

            //// don't run this test when we are using older Xerces as it triggers an XML Parser backwards compatibility issue 
            //// in the xmlsec jar file
            //String AdditionalJar = GetEnvironmentVariable("Additionaljar");
            ////System.out.Println("Having: " + AdditionalJar);
            //Assume.AssumeTrue("Not Running TestSignatureInfo because we are testing with Additionaljar Set to " + AdditionalJar,
            //        AdditionalJar == null || AdditionalJar.Trim().Length == 0);
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void office2007prettyPrintedRels()
        {
            OPCPackage pkg = OPCPackage.Open(testdata.GetFileInfo("office2007prettyPrintedRels.docx"), PackageAccess.READ);
            try
            {
                SignatureConfig sic = new SignatureConfig();
                sic.SetOpcPackage(pkg);
                SignatureInfo si = new SignatureInfo();
                si.SetSignatureConfig(sic);
                bool isValid = si.VerifySignature();
                Assert.IsTrue(isValid);
            }
            finally
            {
                pkg.Close();
            }
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void GetSignerUnsigned()
        {
            String[] testFiles = {
                "hello-world-unsigned.docx",
                "hello-world-unsigned.pptx",
                "hello-world-unsigned.xlsx",
                "hello-world-office-2010-technical-preview-unsigned.docx"
            };

            foreach (String testFile in testFiles)
            {
                OPCPackage pkg = OPCPackage.Open(testdata.GetFileInfo(testFile), PackageAccess.READ);
                SignatureConfig sic = new SignatureConfig();
                sic.SetOpcPackage(pkg);
                SignatureInfo si = new SignatureInfo();
                si.SetSignatureConfig(sic);
                List<X509Certificate> result = new List<X509Certificate>();
                foreach (SignatureInfo.SignaturePart sp in si.GetSignatureParts())
                {
                    if (sp.Validate())
                    {
                        result.Add(sp.GetSigner());
                    }
                }
                pkg.Revert();
                pkg.Close();
                Assert.IsNotNull(result);
                Assert.IsEmpty(result);
            }
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void GetSigner()
        {
            String[] testFiles = {
                "hyperlink-example-signed.docx",
                "hello-world-signed.docx",
                "hello-world-signed.pptx",
                "hello-world-signed.xlsx",
                "hello-world-office-2010-technical-preview.docx",
                "ms-office-2010-signed.docx",
                "ms-office-2010-signed.pptx",
                "ms-office-2010-signed.xlsx",
                "Office2010-SP1-XAdES-X-L.docx",
                "signed.docx",
            };

            foreach (String testFile in testFiles)
            {
                OPCPackage pkg = OPCPackage.Open(testdata.GetFileInfo(testFile), PackageAccess.READ);
                try
                {
                    SignatureConfig sic = new SignatureConfig();
                    sic.SetOpcPackage(pkg);
                    SignatureInfo si = new SignatureInfo();
                    si.SetSignatureConfig(sic);
                    List<X509Certificate> result = new List<X509Certificate>();
                    foreach (SignatureInfo.SignaturePart sp in si.GetSignatureParts())
                    {
                        if (sp.Validate())
                        {
                            result.Add(sp.GetSigner());
                        }
                    }

                    Assert.IsNotNull(result);
                    Assert.AreEqual(1, result.Count, "test-file: " + testFile);
                    X509Certificate signer = result[0];
                    //LOG.Log(POILogger.DEBUG, "signer: " + signer.SubjectX500Principal);

                    bool b = si.VerifySignature();
                    Assert.IsTrue(b, "test-file: " + testFile);
                    pkg.Revert();
                }
                finally
                {
                    pkg.Close();
                }
            }
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void GetMultiSigners()
        {
            String testFile = "hello-world-signed-twice.docx";
            OPCPackage pkg = OPCPackage.Open(testdata.GetFileInfo(testFile), PackageAccess.READ);
            //try {
            //    SignatureConfig sic = new SignatureConfig();
            //    sic.OpcPackage = (/*setter*/pkg);
            //    SignatureInfo si = new SignatureInfo();
            //    si.SignatureConfig = (/*setter*/sic);
            //    List<X509Certificate> result = new List<X509Certificate>();
            //    foreach (SignaturePart sp in si.SignatureParts) {
            //        if (sp.Validate()) {
            //            result.Add(sp.Signer);
            //        }
            //    }

            //    Assert.IsNotNull(result);
            //    Assert.AreEqual("test-file: " + testFile, 2, result.Size());
            //    X509Certificate signer1 = result.Get(0);
            //    X509Certificate signer2 = result.Get(1);
            //    //LOG.Log(POILogger.DEBUG, "signer 1: " + signer1.SubjectX500Principal);
            //    //LOG.Log(POILogger.DEBUG, "signer 2: " + signer2.SubjectX500Principal);

            //    bool b = si.VerifySignature();
            //    Assert.IsTrue("test-file: " + testFile, b);
            //    pkg.Revert();
            //} finally {
            //    pkg.Close();
            //}
            throw new NotImplementedException();
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestSignSpreadsheet()
        {
            String testFile = "hello-world-unsigned.xlsx";
            OPCPackage pkg = OPCPackage.Open(copy(testdata.GetFileInfo(testFile)), PackageAccess.READ_WRITE);
            sign(pkg, "Test", "CN=Test", 1);
            pkg.Close();
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestManipulation()
        {
            //// sign & validate
            //String testFile = "hello-world-unsigned.xlsx";
            //OPCPackage pkg = OPCPackage.Open(copy(testdata.GetFile(testFile)), PackageAccess.READ_WRITE);
            //sign(pkg, "Test", "CN=Test", 1);

            //// manipulate
            //XSSFWorkbook wb = new XSSFWorkbook(pkg);
            //wb.SetSheetName(0, "manipulated");
            //// ... I don't know, why Commit is protected ...
            //Method m = typeof(XSSFWorkbook).GetDeclaredMethod("commit");
            //m.Accessible = (/*setter*/true);
            //m.Invoke(wb);

            //// todo: test a manipulation on a package part, which is not signed
            //// ... maybe in combination with #56164 

            //// validate
            //SignatureConfig sic = new SignatureConfig();
            //sic.OpcPackage = (/*setter*/pkg);
            //SignatureInfo si = new SignatureInfo();
            //si.SignatureConfig = (/*setter*/sic);
            //bool b = si.VerifySignature();
            //Assert.IsFalse("signature should be broken", b);

            //wb.Close();
            throw new NotImplementedException();
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestSignSpreadsheetWithSignatureInfo()
        {
            //InitKeyPair("Test", "CN=Test");
            //String testFile = "hello-world-unsigned.xlsx";
            //OPCPackage pkg = OPCPackage.Open(copy(testdata.GetFile(testFile)), PackageAccess.READ_WRITE);
            //SignatureConfig sic = new SignatureConfig();
            //sic.OpcPackage = (/*setter*/pkg);
            //sic.Key = (/*setter*/keyPair.Private);
            //sic.SigningCertificateChain = (/*setter*/Collections.SingletonList(x509));
            //SignatureInfo si = new SignatureInfo();
            //si.SignatureConfig = (/*setter*/sic);
            //// hash > sha1 doesn't work in excel viewer ...
            //si.ConfirmSignature();
            //List<X509Certificate> result = new List<X509Certificate>();
            //foreach (SignaturePart sp in si.SignatureParts) {
            //    if (sp.Validate()) {
            //        result.Add(sp.Signer);
            //    }
            //}
            //Assert.AreEqual(1, result.Size());
            //pkg.Close();
            throw new NotImplementedException();
        }

        [Ignore("not implemented")]
        public void TestSignEnvelopingDocument()
        {
            //String testFile = "hello-world-unsigned.xlsx";
            //OPCPackage pkg = OPCPackage.Open(copy(testdata.GetFile(testFile)), PackageAccess.READ_WRITE);

            //InitKeyPair("Test", "CN=Test");
            //X509CRL crl = PkiTestUtils.GenerateCrl(x509, keyPair.Private);

            //// Setup
            //SignatureConfig signatureConfig = new SignatureConfig();
            //signatureConfig.OpcPackage=(/*setter*/pkg);
            //signatureConfig.Key=(/*setter*/keyPair.Private);

            ///*
            // * We need at least 2 certificates for the XAdES-C complete certificate
            // * refs construction.
            // */
            //List<X509Certificate> certificateChain = new List<X509Certificate>();
            //certificateChain.Add(x509);
            //certificateChain.Add(x509);
            //signatureConfig.SigningCertificateChain=(/*setter*/certificateChain);

            //signatureConfig.AddSignatureFacet(new EnvelopedSignatureFacet());
            //signatureConfig.AddSignatureFacet(new KeyInfoSignatureFacet());
            //signatureConfig.AddSignatureFacet(new XAdESSignatureFacet());
            //signatureConfig.AddSignatureFacet(new XAdESXLSignatureFacet());

            //// check for internet, no error means it works
            //bool mockTsp = (getAccessError("http://timestamp.comodoca.com/rfc3161", true, 10000) != null);

            //// http://timestamping.edelweb.fr/service/tsp
            //// http://tsa.belgium.be/connect
            //// http://timestamp.comodoca.com/authenticode
            //// http://timestamp.comodoca.com/rfc3161
            //// http://services.globaltrustFinder.com/adss/tsa
            //signatureConfig.TspUrl=(/*setter*/"http://timestamp.comodoca.com/rfc3161");
            //signatureConfig.TspRequestPolicy=(/*setter*/null); // comodoca request fails, if default policy is Set ...
            //signatureConfig.TspOldProtocol=(/*setter*/false);

            ////set proxy info if any
            //String proxy = GetEnvironmentVariable("http_proxy");
            //if (proxy != null && proxy.Trim().Length > 0) {
            //    signatureConfig.ProxyUrl=(/*setter*/proxy);
            //}

            //if (mockTsp) {
            //    //TimeStampService tspService = new TimeStampService(){

            //    //    public byte[] timeStamp(byte[] data, RevocationData revocationData)  {
            //    //        revocationData.AddCRL(crl);
            //    //        return "time-stamp-token".Bytes;
            //    //    }

            //    //    public void SetSignatureConfig(SignatureConfig config) {
            //    //        // empty on purpose
            //    //    }
            //    //};
            //    //signatureConfig.TspService=(/*setter*/tspService);
            //} else {
            //    TimeStampServiceValidator tspValidator = new TimeStampServiceValidator() {

            //        public void validate(List<X509Certificate> certificateChain,
            //        RevocationData revocationData)  {
            //            foreach (X509Certificate certificate in certificateChain) {
            //                LOG.Log(POILogger.DEBUG, "certificate: " + certificate.SubjectX500Principal);
            //                LOG.Log(POILogger.DEBUG, "validity: " + certificate.NotBefore + " - " + certificate.NotAfter);
            //            }
            //        }
            //    };
            //    signatureConfig.TspValidator=(/*setter*/tspValidator);
            //    signatureConfig.TspOldProtocol=(/*setter*/signatureConfig.TspUrl.Contains("edelweb"));
            //}

            //RevocationData revocationData = new RevocationData();
            //revocationData.AddCRL(crl);
            //OCSPResp ocspResp = PkiTestUtils.CreateOcspResp(x509, false,
            //        x509, x509, keyPair.Private, "SHA1withRSA", cal.TimeInMillis);
            //revocationData.AddOCSP(ocspResp.Encoded);

            //RevocationDataService revocationDataService = new RevocationDataService(){

            //    public RevocationData GetRevocationData(List<X509Certificate> certificateChain) {
            //        return revocationData;
            //    }
            //};
            //signatureConfig.RevocationDataService=(/*setter*/revocationDataService);

            //// operate
            //SignatureInfo si = new SignatureInfo();
            //si.SignatureConfig=(/*setter*/signatureConfig);
            //try {
            //    si.ConfirmSignature();
            //} catch (Exception e) {
            //    // only allow a ConnectException because of timeout, we see this in Jenkins from time to time...
            //    if(e.Cause == null) {
            //        throw e;
            //    }
            //    if(!(e.Cause is ConnectException)) {
            //        throw e;
            //    }
            //    Assert.IsTrue("Only allowing ConnectException with 'timed out' as message here, but had: " + e, e.Cause.Message.Contains("timed out"));
            //}

            //// verify
            //Iterator<SignaturePart> spIter = si.SignatureParts.Iterator();
            //Assert.IsTrue(spIter.HasNext());
            //SignaturePart sp = spIter.Next();
            //bool valid = sp.Validate();
            //Assert.IsTrue(valid);

            //SignatureDocument sigDoc = sp.SignatureDocument;
            //String declareNS =
            //    "declare namespace xades='http://uri.etsi.org/01903/v1.3.2#'; "
            //  + "declare namespace ds='http://www.w3.org/2000/09/xmldsig#'; ";

            //String digestValXQuery = declareNS +
            //    "$this/ds:Signature/ds:SignedInfo/ds:Reference";
            //foreach (ReferenceType rt in (ReferenceType[])sigDoc.SelectPath(digestValXQuery)) {
            //    Assert.IsNotNull(rt.DigestValue);
            //    Assert.AreEqual(signatureConfig.DigestMethodUri, rt.DigestMethod.Algorithm);
            //}

            //String certDigestXQuery = declareNS +
            //    "$this//xades:SigningCertificate/xades:Cert/xades:CertDigest";
            //XmlObject xoList[] = sigDoc.SelectPath(certDigestXQuery);
            //Assert.AreEqual(xoList.Length, 1);
            //DigestAlgAndValueType certDigest = (DigestAlgAndValueType)xoList[0];
            //Assert.IsNotNull(certDigest.DigestValue);

            //String qualPropXQuery = declareNS +
            //    "$this/ds:Signature/ds:Object/xades:QualifyingProperties";
            //xoList = sigDoc.SelectPath(qualPropXQuery);
            //Assert.AreEqual(xoList.Length, 1);
            //QualifyingPropertiesType qualProp = (QualifyingPropertiesType)xoList[0];
            //bool qualPropXsdOk = qualProp.Validate();
            //Assert.IsTrue(qualPropXsdOk);

            //pkg.Close();
            throw new NotImplementedException();
        }

        public static String GetAccessError(String destinationUrl, bool fireRequest, int timeout)
        {
            //URL url;
            //try {
            //    url = new URL(destinationUrl);
            //} catch (MalformedURLException e) {
            //    throw new ArgumentException("Invalid destination URL", e);
            //}

            //HttpURLConnection conn = null;
            //try {
            //    conn = (HttpURLConnection) url.OpenConnection();

            //    // Set specified timeout if non-zero
            //    if(timeout != 0) {
            //        conn.ConnectTimeout=(/*setter*/timeout);
            //        conn.ReadTimeout=(/*setter*/timeout);
            //    }

            //    conn.DoOutput=(/*setter*/false);
            //    conn.DoInput=(/*setter*/true);

            //    /* if connecting is not possible this will throw a connection refused exception */
            //    conn.Connect();

            //    if (fireRequest) {
            //        InputStream is1 = null;
            //        try {
            //            is1 = conn.InputStream;
            //        } finally {
            //            IOUtils.CloseQuietly(is1);
            //        }

            //    }
            //    /* if connecting is possible we return true here */
            //    return null;

            //} catch (IOException e) {
            //    /* exception is thrown -> server not available */
            //    return e.Class.Name + ": " + e.Message;
            //} finally {
            //    if (conn != null) {
            //        conn.Disconnect();
            //    }
            //}
            throw new NotImplementedException();
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestCertChain()
        {
            //KeyStore keystore = KeyStore.GetInstance("PKCS12");
            //String password = "test";
            //InputStream is1 = testdata.OpenResourceAsStream("chaintest.pfx");
            //keystore.Load(is, password.ToCharArray());
            //is1.Close();

            //Key key = keystore.GetKey("poitest", password.ToCharArray());
            //Certificate chainList[] = keystore.GetCertificateChain("poitest");
            //List<X509Certificate> certChain = new List<X509Certificate>();
            //foreach (Certificate c in chainList) {
            //    certChain.Add((X509Certificate)c);
            //}
            //x509 = certChain.Get(0);
            //keyPair = new KeyPair(x509.PublicKey, (PrivateKey)key);

            //String testFile = "hello-world-unsigned.xlsx";
            //OPCPackage pkg = OPCPackage.Open(copy(testdata.GetFile(testFile)), PackageAccess.READ_WRITE);

            //SignatureConfig signatureConfig = new SignatureConfig();
            //signatureConfig.Key = (/*setter*/keyPair.Private);
            //signatureConfig.SigningCertificateChain = (/*setter*/certChain);
            //Calendar cal = Calendar.Instance;
            //cal.Set(2007, 7, 1);
            //signatureConfig.ExecutionTime = (/*setter*/cal.Time);
            //signatureConfig.DigestAlgo = (/*setter*/HashAlgorithm.sha1);
            //signatureConfig.OpcPackage = (/*setter*/pkg);

            //SignatureInfo si = new SignatureInfo();
            //si.SignatureConfig = (/*setter*/signatureConfig);

            //si.ConfirmSignature();

            //foreach (SignaturePart sp in si.SignatureParts) {
            //    Assert.IsTrue("Could not validate", sp.Validate());
            //    X509Certificate signer = sp.Signer;
            //    Assert.IsNotNull("signer undefined?!", signer);
            //    List<X509Certificate> certChainRes = sp.CertChain;
            //    Assert.AreEqual(3, certChainRes.Size());
            //}

            //pkg.Close();
            throw new NotImplementedException();
        }

        [Ignore("not implemented")]
        public void TestNonSha1()
        {
            //String testFile = "hello-world-unsigned.xlsx";
            //InitKeyPair("Test", "CN=Test");

            //SignatureConfig signatureConfig = new SignatureConfig();
            //signatureConfig.Key = (/*setter*/keyPair.Private);
            //signatureConfig.SigningCertificateChain = (/*setter*/Collections.SingletonList(x509));

            //HashAlgorithm testAlgo[] = { HashAlgorithm.sha224, HashAlgorithm.sha256
            //    , HashAlgorithm.sha384, HashAlgorithm.sha512, HashAlgorithm.ripemd160 };

            //foreach (HashAlgorithm ha in testAlgo) {
            //    OPCPackage pkg = null;
            //    try {
            //        signatureConfig.DigestAlgo = (/*setter*/ha);
            //        pkg = OPCPackage.Open(copy(testdata.GetFile(testFile)), PackageAccess.READ_WRITE);
            //        signatureConfig.OpcPackage = (/*setter*/pkg);

            //        SignatureInfo si = new SignatureInfo();
            //        si.SignatureConfig = (/*setter*/signatureConfig);

            //        si.ConfirmSignature();
            //        bool b = si.VerifySignature();
            //        Assert.IsTrue("Signature not correctly calculated for " + ha, b);
            //    } finally {
            //        if (pkg != null) pkg.Close();
            //    }
            //}
            throw new NotImplementedException();
        }

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestMultiSign()
        {
            //initKeyPair("KeyA", "CN=KeyA");
            //KeyPair keyPairA = keyPair;
            //X509Certificate x509A = x509;
            //initKeyPair("KeyB", "CN=KeyB");
            //KeyPair keyPairB = keyPair;
            //X509Certificate x509B = x509;

            //File tpl = copy(testdata.GetFile("bug58630.xlsx"));
            //OPCPackage pkg = OPCPackage.open(tpl);
            //SignatureConfig signatureConfig = new SignatureConfig();


        }
        private void sign(OPCPackage pkgCopy, String alias, String signerDn, int signerCount)
        {
            throw new NotImplementedException();
            //InitKeyPair(alias, signerDn);

            //SignatureConfig signatureConfig = new SignatureConfig();
            //signatureConfig.Key=(/*setter*/keyPair.Private);
            //signatureConfig.SigningCertificateChain=(/*setter*/Collections.SingletonList(x509));
            //signatureConfig.ExecutionTime=(/*setter*/cal.Time);
            //signatureConfig.DigestAlgo=(/*setter*/HashAlgorithm.sha1);
            //signatureConfig.OpcPackage=(/*setter*/pkgCopy);

            //SignatureInfo si = new SignatureInfo();
            //si.SignatureConfig=(/*setter*/signatureConfig);

            //Document document = DocumentHelper.CreateDocument();

            //// operate
            //DigestInfo digestInfo = si.PreSign(document, null);

            //// verify
            //Assert.IsNotNull(digestInfo);
            //LOG.Log(POILogger.DEBUG, "digest algo: " + digestInfo.HashAlgo);
            //LOG.Log(POILogger.DEBUG, "digest description: " + digestInfo.description);
            //Assert.AreEqual("Office OpenXML Document", digestInfo.description);
            //Assert.IsNotNull(digestInfo.HashAlgo);
            //Assert.IsNotNull(digestInfo.digestValue);

            //// Setup: key material, signature value
            //byte[] signatureValue = si.SignDigest(digestInfo.digestValue);

            //// operate: postSign
            //si.PostSign(document, signatureValue);

            //// verify: signature
            //si.SignatureConfig.OpcPackage=(/*setter*/pkgCopy);
            //List<X509Certificate> result = new List<X509Certificate>();
            //foreach (SignaturePart sp in si.SignatureParts) {
            //    if (sp.Validate()) {
            //        result.Add(sp.Signer);
            //    }
            //}
            //Assert.AreEqual(signerCount, result.Size());
        }

        private void InitKeyPair(String alias, String subjectDN)
        {
            throw new NotImplementedException();
            //char password[] = "test".ToCharArray();
            //File file = new File("build/test.pfx");

            //KeyStore keystore = KeyStore.GetInstance("PKCS12");

            //if (file.Exists()) {
            //    FileInputStream fis = new FileInputStream(file);
            //    keystore.Load(fis, password);
            //    fis.Close();
            //} else {
            //    keystore.Load(null, password);
            //}

            //if (keystore.IsKeyEntry(alias)) {
            //    Key key = keystore.GetKey(alias, password);
            //    x509 = (X509Certificate)keystore.GetCertificate(alias);
            //    keyPair = new KeyPair(x509.PublicKey, (PrivateKey)key);
            //} else {
            //    keyPair = PkiTestUtils.GenerateKeyPair();
            //    Calendar cal = Calendar.Instance;
            //    Date notBefore = cal.Time;
            //    cal.Add(Calendar.YEAR, 1);
            //    Date notAfter = cal.Time;
            //    KeyUsage keyUsage = new KeyUsage(KeyUsage.digitalSignature);

            //    x509 = PkiTestUtils.GenerateCertificate(keyPair.Public, subjectDN
            //        , notBefore, notAfter, null, keyPair.Private, true, 0, null, null, keyUsage);

            //    keystore.KeyEntry=(/*setter*/alias, keyPair.Private, password, new Certificate[]{x509});
            //    FileOutputStream fos = new FileOutputStream(file);
            //    keystore.Store(fos, password);
            //    fos.Close();
            //}
        }

        private static FileInfo copy(FileInfo input)
        {
            String extension = input.Name.Replace(".*?(\\.[^.]+)?$", "$1");
            if (extension == null || "".Equals(extension))
                extension = ".zip";
            FileInfo tmpFile = new FileInfo("build" + Path.DirectorySeparatorChar + "sigtest" + extension);
            throw new NotImplementedException();
            //FileStream fos = tmpFile.Create();// FileOutputStream(tmpFile);
            //FileStream fis = input.Create();// new FileInputStream(input);
            //IOUtils.Copy(fis, fos);
            //fis.Close();
            //fos.Close();
            //return tmpFile;
        }

    }

}
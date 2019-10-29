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

namespace NPOI.POIFS.Crypt.Dsig.Services
{
    using NPOI.POIFS.Crypt.Dsig;
    using NPOI.Util;
    using Org.BouncyCastle.Asn1;
    using Org.BouncyCastle.Asn1.Cmp;
    using Org.BouncyCastle.Asn1.Nist;
    using Org.BouncyCastle.Asn1.X509;
    using Org.BouncyCastle.Cms;
    using Org.BouncyCastle.Tsp;
    using Org.BouncyCastle.X509;
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;


    /**
     * A TSP time-stamp service implementation.
     * 
     * @author Frank Cornelis
     * 
     */
    public class TSPTimeStampService : ITimeStampService
    {

        //private static POILogger LOG = POILogFactory.GetLogger(typeof(TSPTimeStampService));

        private SignatureConfig signatureConfig;

        /**
         * Maps the digest algorithm to corresponding OID value.
         */
        public DerObjectIdentifier mapDigestAlgoToOID(HashAlgorithm digestAlgo)
        {
            switch (digestAlgo.jceId)
            {
                case "sha1": return X509ObjectIdentifiers.IdSha1;
                case "sha256": return NistObjectIdentifiers.IdSha256;
                case "sha384": return NistObjectIdentifiers.IdSha384;
                case "sha512": return NistObjectIdentifiers.IdSha512;
                default:
                    throw new ArgumentException("unsupported digest algo: " + digestAlgo);
            }
        }

        public byte[] TimeStamp(byte[] data, RevocationData revocationData)
        {
            // digest the message
            MessageDigest messageDigest = CryptoFunctions.GetMessageDigest(signatureConfig.GetTspDigestAlgo());
            byte[] digest = messageDigest.Digest(data);
            throw new NotImplementedException();
            //// generate the TSP request
            //Bigint nonce = new Bigint(128, new Random());
            //TimeStampRequestGenerator requestGenerator = new TimeStampRequestGenerator();
            //requestGenerator.SetCertReq(true);
            //String requestPolicy = signatureConfig.GetTspRequestPolicy();
            //if (requestPolicy != null)
            //{
            //    requestGenerator.SetReqPolicy(new ASN1ObjectIdentifier(requestPolicy));
            //}
            //ASN1ObjectIdentifier digestAlgoOid = mapDigestAlgoToOID(signatureConfig.GetTspDigestAlgo());
            //TimeStampRequest request = requestGenerator.Generate(digestAlgoOid, digest, nonce);
            //byte[] encodedRequest = request.GetEncoded();

            //// create the HTTP POST request
            //Proxy proxy = Proxy.NO_PROXY;
            //if (signatureConfig.ProxyUrl != null)
            //{
            //    URL proxyUrl = new URL(signatureConfig.ProxyUrl);
            //    String host = proxyUrl.Host;
            //    int port = proxyUrl.Port;
            //    proxy = new Proxy(Proxy.Type.HTTP, new InetSocketAddress(host, (port == -1 ? 80 : port)));
            //}

            //HttpURLConnection huc = (HttpURLConnection)new URL(signatureConfig.TspUrl).OpenConnection(proxy);

            //if (signatureConfig.GetTspUser() != null)
            //{
            //    String userPassword = signatureConfig.GetTspUser() + ":" + signatureConfig.GetTspPass();
            //    String encoding = DatatypeConverter.PrintBase64Binary(userPassword.GetBytes(Charset.ForName("iso-8859-1")));
            //    huc.RequestProperty = (/*setter*/"Authorization", "Basic " + encoding);
            //}

            //huc.RequestMethod = (/*setter*/"POST");
            //huc.ConnectTimeout = (/*setter*/20000);
            //huc.ReadTimeout = (/*setter*/20000);
            //huc.DoOutput = (/*setter*/true); // also Sets method to POST.
            //huc.RequestProperty = (/*setter*/"User-Agent", signatureConfig.UserAgent);
            //huc.RequestProperty = (/*setter*/"Content-Type", signatureConfig.IsTspOldProtocol()
            //    ? "application/timestamp-request"
            //    : "application/timestamp-query"); // "; charset=ISO-8859-1");

            //OutputStream hucOut = huc.OutputStream;
            //hucOut.Write(encodedRequest);

            //// invoke TSP service
            //huc.Connect();

            //int statusCode = huc.ResponseCode;
            //if (statusCode != 200)
            //{
            //    //LOG.Log(POILogger.ERROR, "Error contacting TSP server ", signatureConfig.TspUrl);
            //    throw new IOException("Error contacting TSP server " + signatureConfig.TspUrl);
            //}

            //// HTTP input validation
            //String contentType = huc.GetHeaderField("Content-Type");
            //if (null == contentType)
            //{
            //    throw new Exception("missing Content-Type header");
            //}

            //MemoryStream bos = new MemoryStream();
            //IOUtils.Copy(huc.InputStream, bos);
            ////LOG.Log(POILogger.DEBUG, "response content: ", bos.ToString());

            //if (!contentType.StartsWith(signatureConfig.IsTspOldProtocol()
            //    ? "application/timestamp-response"
            //    : "application/timestamp-reply"
            //))
            //{
            //    throw new Exception("invalid Content-Type: " + contentType);
            //}

            //if (bos.Length == 0)
            //{
            //    throw new Exception("Content-Length is zero");
            //}

            //// TSP response parsing and validation
            //TimeStampResponse timeStampResponse = new TimeStampResponse(bos.ToArray());
            //timeStampResponse.Validate(request);

            //if (0 != timeStampResponse.Status)
            //{
            //    //LOG.Log(POILogger.DEBUG, "status: " + timeStampResponse.Status);
            //    //LOG.Log(POILogger.DEBUG, "status string: " + timeStampResponse.StatusString);
            //    PkiFailureInfo failInfo = timeStampResponse.GetFailInfo();
            //    if (null != failInfo)
            //    {
            //        //LOG.Log(POILogger.DEBUG, "fail info int value: " + failInfo.IntValue());
            //        if (/*PKIFailureInfo.unacceptedPolicy*/(1 << 8) == failInfo.IntValue)
            //        {
            //            //LOG.Log(POILogger.DEBUG, "unaccepted policy");
            //        }
            //    }
            //    throw new Exception("timestamp response status != 0: "
            //            + timeStampResponse.Status);
            //}
            //TimeStampToken timeStampToken = timeStampResponse.TimeStampToken;
            //SignerID signerId = timeStampToken.SignerID;
            //Bigint signerCertSerialNumber = signerId.SerialNumber;
            //X500Name signerCertIssuer = signerId.Issuer;
            ////LOG.Log(POILogger.DEBUG, "signer cert serial number: " + signerCertSerialNumber);
            ////LOG.Log(POILogger.DEBUG, "signer cert issuer: " + signerCertIssuer);

            //// TSP signer certificates retrieval
            //Collection<X509CertificateHolder> certificates = timeStampToken.GetCertificates().GetMatches(null);

            //X509CertificateHolder signerCert = null;
            //Dictionary<X500Name, X509CertificateHolder> certificateMap = new Dictionary<X500Name, X509CertificateHolder>();
            //foreach (X509CertificateHolder certificate in certificates)
            //{
            //    if (signerCertIssuer.Equals(certificate.Issuer)
            //        && signerCertSerialNumber.Equals(certificate.SerialNumber))
            //    {
            //        signerCert = certificate;
            //    }
            //    certificateMap.Add(certificate.Subject, certificate);
            //}

            //// TSP signer cert path building
            //if (signerCert == null)
            //{
            //    throw new Exception("TSP response token has no signer certificate");
            //}
            //List<X509Certificate> tspCertificateChain = new List<X509Certificate>();
            //JcaX509CertificateConverter x509Converter = new JcaX509CertificateConverter();
            //x509Converter.Provider = (/*setter*/"BC");
            //X509CertificateHolder certificate = signerCert;
            //do
            //{
            //    //LOG.Log(POILogger.DEBUG, "Adding to certificate chain: " + certificate.Subject);
            //    tspCertificateChain.Add(x509converter.GetCertificate(certificate));
            //    if (certificate.Subject.Equals(certificate.Issuer))
            //    {
            //        break;
            //    }
            //    certificate = certificateMap.Get(certificate.Issuer);
            //} while (null != certificate);

            //// verify TSP signer signature
            //X509CertificateHolder holder = new X509CertificateHolder(tspCertificateChain.Get(0).Encoded);
            //DefaultCMSSignatureAlgorithmNameGenerator nameGen = new DefaultCMSSignatureAlgorithmNameGenerator();
            //DefaultSignatureAlgorithmIdentifierFinder sigAlgoFinder = new DefaultSignatureAlgorithmIdentifierFinder();
            //DefaultDigestAlgorithmIdentifierFinder hashAlgoFinder = new DefaultDigestAlgorithmIdentifierFinder();
            //BcDigestCalculatorProvider calculator = new BcDigestCalculatorProvider();
            //BcRSASignerInfoVerifierBuilder verifierBuilder = new BcRSASignerInfoVerifierBuilder(nameGen, sigAlgoFinder, hashAlgoFinder, calculator);
            //SignerInformationVerifier verifier = verifierBuilder.Build(holder);

            //timeStampToken.Validate(verifier);

            //// verify TSP signer certificate
            //if (signatureConfig.GetTspValidator() != null)
            //{
            //    signatureConfig.GetTspValidator().Validate(tspCertificateChain, revocationData);
            //}

            ////LOG.Log(POILogger.DEBUG, "time-stamp token time: "
            ////        + timeStampToken.TimeStampInfo.GenTime);

            //byte[] timestamp = timeStampToken.GetEncoded();
            //return timestamp;
        }

        public void SetSignatureConfig(SignatureConfig signatureConfig)
        {
            this.signatureConfig = signatureConfig;
        }
    }
}
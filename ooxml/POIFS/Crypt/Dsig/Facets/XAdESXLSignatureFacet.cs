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

namespace NPOI.POIFS.Crypt.Dsig.Facets
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;

    /**
     * XAdES-X-L v1.4.1 signature facet. This signature facet implementation will
     * upgrade a given XAdES-BES/EPES signature to XAdES-X-L.
     * 
     * We don't inherit from XAdESSignatureFacet as we also want to be able to use
     * this facet out of the context of a signature creation. This signature facet
     * assumes that the signature is already XAdES-BES/EPES compliant.
     * 
     * This implementation has been tested against an implementation that
     * participated multiple ETSI XAdES plugtests.
     * 
     * @author Frank Cornelis
     * @see XAdESSignatureFacet
     */
    public class XAdESXLSignatureFacet : SignatureFacet {

        //private static POILogger LOG = POILogFactory.GetLogger(typeof(XAdESXLSignatureFacet));

        //private CertificateFactory certificateFactory;

        public XAdESXLSignatureFacet() {
            //try {
            //    this.certificateFactory = CertificateFactory.GetInstance("X.509");
            //} catch (CertificateException e) {
            //    throw new Exception("X509 JCA error: " + e.Message, e);
            //}
            throw new NotImplementedException();
        }


        public override void postSign(XmlDocument document) {
            throw new NotImplementedException();
            //LOG.Log(POILogger.DEBUG, "XAdES-X-L post sign phase");

            //QualifyingPropertiesDocument qualDoc = null;
            //QualifyingPropertiesType qualProps = null;

            //// check for XAdES-BES
            //NodeList qualNl = document.GetElementsByTagNameNS(XADES_132_NS, "QualifyingProperties");
            //if (qualNl.Length == 1) {
            //    try {
            //        qualDoc = QualifyingPropertiesDocument.Factory.Parse(qualNl.Item(0));
            //    } catch (XmlException e) {
            //        throw new MarshalException(e);
            //    }
            //    qualProps = qualDoc.QualifyingProperties;
            //} else {
            //    throw new MarshalException("no XAdES-BES extension present");
            //}

            //// create basic XML Container structure
            //UnsignedPropertiesType unsignedProps = qualProps.UnsignedProperties;
            //if (unsignedProps == null) {
            //    unsignedProps = qualProps.AddNewUnsignedProperties();
            //}
            //UnsignedSignaturePropertiesType unsignedSigProps = unsignedProps.UnsignedSignatureProperties;
            //if (unsignedSigProps == null) {
            //    unsignedSigProps = unsignedProps.AddNewUnsignedSignatureProperties();
            //}


            //// create the XAdES-T time-stamp
            //NodeList nlSigVal = document.GetElementsByTagNameNS(XML_DIGSIG_NS, "SignatureValue");
            //if (nlSigVal.Length != 1) {
            //    throw new ArgumentException("SignatureValue is not Set.");
            //}

            //RevocationData tsaRevocationDataXadesT = new RevocationData();
            //LOG.Log(POILogger.DEBUG, "creating XAdES-T time-stamp");
            //XAdESTimeStampType signatureTimeStamp = CreateXAdESTimeStamp
            //    (Collections.SingletonList(nlSigVal.Item(0)), tsaRevocationDataXadesT);

            //// marshal the XAdES-T extension
            //unsignedSigProps.AddNewSignatureTimeStamp().Set(signatureTimeStamp);

            //// xadesv141::TimeStampValidationData
            //if (tsaRevocationDataXadesT.HasRevocationDataEntries()) {
            //    ValidationDataType validationData = CreateValidationData(tsaRevocationDataXadesT);
            //    insertXChild(unsignedSigProps, validationData);
            //}

            //if (signatureConfig.RevocationDataService == null) {
            //    /*
            //     * Without revocation data service we cannot construct the XAdES-C
            //     * extension.
            //     */
            //    return;
            //}

            //// XAdES-C: complete certificate refs
            //CompleteCertificateRefsType completeCertificateRefs =
            //    unsignedSigProps.AddNewCompleteCertificateRefs();

            //CertIDListType certIdList = completeCertificateRefs.AddNewCertRefs();
            ///*
            // * We skip the signing certificate itself according to section
            // * 4.4.3.2 of the XAdES 1.4.1 specification.
            // */
            //List<X509Certificate> certChain = signatureConfig.SigningCertificateChain;
            //int chainSize = certChain.Size();
            //if (chainSize > 1) {
            //    foreach (X509Certificate cert in certChain.SubList(1, chainSize)) {
            //        CertIDType certId = certIdList.AddNewCert();
            //        XAdESSignatureFacet.CertID = (/*setter*/certId, signatureConfig, false, cert);
            //    }
            //}

            //// XAdES-C: complete revocation refs
            //CompleteRevocationRefsType completeRevocationRefs =
            //    unsignedSigProps.AddNewCompleteRevocationRefs();
            //RevocationData revocationData = signatureConfig.RevocationDataService
            //    .GetRevocationData(certChain);
            //if (revocationData.HasCRLs()) {
            //    CRLRefsType crlRefs = completeRevocationRefs.AddNewCRLRefs();
            //    completeRevocationRefs.CRLRefs = (/*setter*/crlRefs);

            //    foreach (byte[] encodedCrl in revocationData.CRLs) {
            //        CRLRefType crlRef = crlRefs.AddNewCRLRef();
            //        X509CRL crl;
            //        try {
            //            crl = (X509CRL)this.certificateFactory
            //                    .generateCRL(new MemoryStream(encodedCrl));
            //        } catch (CRLException e) {
            //            throw new Exception("CRL parse error: "
            //                    + e.Message, e);
            //        }

            //        CRLIdentifierType crlIdentifier = crlRef.AddNewCRLIdentifier();
            //        String issuerName = crl.IssuerDN.Name.Replace(",", ", ");
            //        crlIdentifier.Issuer = (/*setter*/issuerName);
            //        Calendar cal = Calendar.Instance;
            //        cal.Time = (/*setter*/crl.ThisUpdate);
            //        crlIdentifier.IssueTime = (/*setter*/cal);
            //        crlIdentifier.Number = (/*setter*/getCrlNumber(crl));

            //        DigestAlgAndValueType digestAlgAndValue = crlRef.AddNewDigestAlgAndValue();
            //        XAdESSignatureFacet.DigestAlgAndValue = (/*setter*/digestAlgAndValue, encodedCrl, signatureConfig.DigestAlgo);
            //    }
            //}
            //if (revocationData.HasOCSPs()) {
            //    OCSPRefsType ocspRefs = completeRevocationRefs.AddNewOCSPRefs();
            //    foreach (byte[] ocsp in revocationData.OCSPs) {
            //        try {
            //            OCSPRefType ocspRef = ocspRefs.AddNewOCSPRef();

            //            DigestAlgAndValueType digestAlgAndValue = ocspRef.AddNewDigestAlgAndValue();
            //            XAdESSignatureFacet.DigestAlgAndValue = (/*setter*/digestAlgAndValue, ocsp, signatureConfig.DigestAlgo);

            //            OCSPIdentifierType ocspIdentifier = ocspRef.AddNewOCSPIdentifier();

            //            OCSPResp ocspResp = new OCSPResp(ocsp);

            //            BasicOCSPResp basicOcspResp = (BasicOCSPResp)ocspResp.ResponseObject;

            //            Calendar cal = Calendar.Instance;
            //            cal.Time = (/*setter*/basicOcspResp.ProducedAt);
            //            ocspIdentifier.ProducedAt = (/*setter*/cal);

            //            ResponderIDType responderId = ocspIdentifier.AddNewResponderID();

            //            RespID respId = basicOcspResp.ResponderId;
            //            ResponderID ocspResponderId = respId.ToASN1Primitive();
            //            DERTaggedObject derTaggedObject = (DERTaggedObject)ocspResponderId.ToASN1Primitive();
            //            if (2 == derTaggedObject.TagNo) {
            //                ASN1OctetString keyHashOctetString = (ASN1OctetString)derTaggedObject.Object;
            //                byte key[] = keyHashOctetString.Octets;
            //                responderId.ByKey = (/*setter*/key);
            //            } else {
            //                X500Name name = X500Name.GetInstance(derTaggedObject.Object);
            //                String nameStr = name.ToString();
            //                responderId.ByName = (/*setter*/nameStr);
            //            }
            //        } catch (Exception e) {
            //            throw new Exception("OCSP decoding error: " + e.Message, e);
            //        }
            //    }
            //}

            //// marshal XAdES-C

            //// XAdES-X Type 1 timestamp
            //List<Node> timeStampNodesXadesX1 = new List<Node>();
            //timeStampNodesXadesX1.Add(nlSigVal.Item(0));
            //timeStampNodesXadesX1.Add(signatureTimeStamp.DomNode);
            //timeStampNodesXadesX1.Add(completeCertificateRefs.DomNode);
            //timeStampNodesXadesX1.Add(completeRevocationRefs.DomNode);

            //RevocationData tsaRevocationDataXadesX1 = new RevocationData();
            //LOG.Log(POILogger.DEBUG, "creating XAdES-X time-stamp");
            //XAdESTimeStampType timeStampXadesX1 = CreateXAdESTimeStamp
            //    (timeStampNodesXadesX1, tsaRevocationDataXadesX1);
            //if (tsaRevocationDataXadesX1.HasRevocationDataEntries()) {
            //    ValidationDataType timeStampXadesX1ValidationData = CreateValidationData(tsaRevocationDataXadesX1);
            //    insertXChild(unsignedSigProps, timeStampXadesX1ValidationData);
            //}

            //// marshal XAdES-X
            //unsignedSigProps.AddNewSigAndRefsTimeStamp().Set(timeStampXadesX1);

            //// XAdES-X-L
            //CertificateValuesType certificateValues = unsignedSigProps.AddNewCertificateValues();
            //foreach (X509Certificate certificate in certChain) {
            //    EncapsulatedPKIDataType encapsulatedPKIDataType = certificateValues.AddNewEncapsulatedX509Certificate();
            //    try {
            //        encapsulatedPKIDataType.ByteArrayValue = (/*setter*/certificate.Encoded);
            //    } catch (CertificateEncodingException e) {
            //        throw new Exception("certificate encoding error: " + e.Message, e);
            //    }
            //}

            //RevocationValuesType revocationValues = unsignedSigProps.AddNewRevocationValues();
            //CreateRevocationValues(revocationValues, revocationData);

            //// marshal XAdES-X-L
            //Node n = document.ImportNode(qualProps.DomNode, true);
            //qualNl.Item(0).ParentNode.ReplaceChild(n, qualNl.Item(0));
        }

        public static byte[] GetC14nValue(List<XmlNode> nodeList, String c14nAlgoId) {
            throw new NotImplementedException();
            //MemoryStream c14nValue = new MemoryStream();
            //try {
            //    foreach (Node node in nodeList) {
            //        /*
            //         * Re-Initialize the c14n else the namespaces will Get cached
            //         * and will be missing from the c14n resulting nodes.
            //         */
            //        Canonicalizer c14n = Canonicalizer.GetInstance(c14nAlgoId);
            //        c14nValue.Write(c14n.CanonicalizeSubtree(node));
            //    }
            //} catch (Exception e) {
            //    throw e;
            //} catch (Exception e) {
            //    throw new Exception("c14n error: " + e.Message, e);
            //}
            //return c14nValue.ToByteArray();
        }

        //private BigInteger getCrlNumber(X509CRL crl)
        //{
        //    byte[] crlNumberExtensionValue = crl.getExtensionValue(Extension.cRLNumber.getId());
        //    if (null == crlNumberExtensionValue)
        //    {
        //        return null;
        //    }

        //    try
        //    {
        //        ASN1InputStream asn1IS1 = null, asn1IS2 = null;
        //        try
        //        {
        //            asn1IS1 = new ASN1InputStream(crlNumberExtensionValue);
        //            ASN1OctetString octetString = (ASN1OctetString)asn1IS1.readObject();
        //            byte[] octets = octetString.getOctets();
        //            asn1IS2 = new ASN1InputStream(octets);
        //            ASN1Integer integer = (ASN1Integer)asn1IS2.readObject();
        //            return integer.getPositiveValue();
        //        }
        //        finally
        //        {
        //            asn1IS2.close();
        //            asn1IS1.close();
        //        }
        //    }
        //    catch (IOException e)
        //    {
        //        throw new RuntimeException("I/O error: " + e.getMessage(), e);
        //    }
        //}

        //private XAdESTimeStampType CreateXAdESTimeStamp(
        //        List<Node> nodeList,
        //        RevocationData revocationData) {
        //    byte[] c14nSignatureValueElement = GetC14nValue(nodeList, signatureConfig.XadesCanonicalizationMethod);

        //    return CreateXAdESTimeStamp(c14nSignatureValueElement, revocationData);
        //}

        //private XAdESTimeStampType CreateXAdESTimeStamp(byte[] data, RevocationData revocationData) {
        //    // create the time-stamp
        //    byte[] timeStampToken;
        //    try {
        //        timeStampToken = signatureConfig.TspService.TimeStamp(data, revocationData);
        //    } catch (Exception e) {
        //        throw new Exception("error while creating a time-stamp: "
        //                + e.Message, e);
        //    }

        //    // create a XAdES time-stamp Container
        //    XAdESTimeStampType xadesTimeStamp = XAdESTimeStampType.Factory.NewInstance();
        //    xadesTimeStamp.Id = (/*setter*/"time-stamp-" + UUID.RandomUUID().ToString());
        //    CanonicalizationMethodType c14nMethod = xadesTimeStamp.AddNewCanonicalizationMethod();
        //    c14nMethod.Algorithm = (/*setter*/signatureConfig.XadesCanonicalizationMethod);

        //    // embed the time-stamp
        //    EncapsulatedPKIDataType encapsulatedTimeStamp = xadesTimeStamp.AddNewEncapsulatedTimeStamp();
        //    encapsulatedTimeStamp.ByteArrayValue = (/*setter*/timeStampToken);
        //    encapsulatedTimeStamp.Id = (/*setter*/"time-stamp-token-" + UUID.RandomUUID().ToString());

        //    return xadesTimeStamp;
        //}

        //private ValidationDataType CreateValidationData(
        //        RevocationData revocationData) {
        //    ValidationDataType validationData = ValidationDataType.Factory.NewInstance();
        //    RevocationValuesType revocationValues = validationData.AddNewRevocationValues();
        //    CreateRevocationValues(revocationValues, revocationData);
        //    return validationData;
        //}

        //private void CreateRevocationValues(
        //        RevocationValuesType revocationValues, RevocationData revocationData) {
        //    if (revocationData.HasCRLs()) {
        //        CRLValuesType crlValues = revocationValues.AddNewCRLValues();
        //        foreach (byte[] crl in revocationData.CRLs) {
        //            EncapsulatedPKIDataType encapsulatedCrlValue = crlValues.AddNewEncapsulatedCRLValue();
        //            encapsulatedCrlValue.ByteArrayValue = (/*setter*/crl);
        //        }
        //    }
        //    if (revocationData.HasOCSPs()) {
        //        OCSPValuesType ocspValues = revocationValues.AddNewOCSPValues();
        //        foreach (byte[] ocsp in revocationData.OCSPs) {
        //            EncapsulatedPKIDataType encapsulatedOcspValue = ocspValues.AddNewEncapsulatedOCSPValue();
        //            encapsulatedOcspValue.ByteArrayValue = (/*setter*/ocsp);
        //        }
        //    }
        //}
    }

}
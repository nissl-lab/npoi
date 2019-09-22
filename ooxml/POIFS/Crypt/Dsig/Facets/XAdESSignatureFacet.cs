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
    using System.Security.Cryptography.Xml;
    using System.Xml;
    
    using NPOI.POIFS.Crypt;


    /**
     * XAdES Signature Facet. Implements XAdES v1.4.1 which is compatible with XAdES
     * v1.3.2. The implemented XAdES format is XAdES-BES/EPES. It's up to another
     * part of the signature service to upgrade the XAdES-BES to a XAdES-X-L.
     * 
     * This implementation has been tested against an implementation that
     * participated multiple ETSI XAdES plugtests.
     * 
     * @author Frank Cornelis
     * @see <a href="http://en.wikipedia.org/wiki/XAdES">XAdES</a>
     * 
     */
    public class XAdESSignatureFacet : SignatureFacet
    {

        private static String XADES_TYPE = "http://uri.etsi.org/01903#SignedProperties";

        private Dictionary<String, String> dataObjectFormatMimeTypes = new Dictionary<String, String>();



        public void preSign(
              XmlDocument document
            , List<Reference> references
            , List<XmlNode> objects)
        {
            ////LOG.Log(POILogger.DEBUG, "preSign");

            //// QualifyingProperties
            //QualifyingPropertiesDocument qualDoc = QualifyingPropertiesDocument.Factory.NewInstance();
            //QualifyingPropertiesType qualifyingProperties = qualDoc.AddNewQualifyingProperties();
            //qualifyingProperties.Target = (/*setter*/"#" + signatureConfig.PackageSignatureId);

            //// SignedProperties
            //SignedPropertiesType signedProperties = qualifyingProperties.AddNewSignedProperties();
            //signedProperties.Id = (/*setter*/signatureConfig.XadesSignatureId);

            //// SignedSignatureProperties
            //SignedSignaturePropertiesType signedSignatureProperties = signedProperties.AddNewSignedSignatureProperties();

            //// SigningTime
            //Calendar xmlGregorianCalendar = Calendar.Instance;
            //xmlGregorianCalendar.TimeZone = (/*setter*/TimeZone.GetTimeZone("Z"));
            //xmlGregorianCalendar.Time = (/*setter*/signatureConfig.ExecutionTime);
            //xmlGregorianCalendar.Clear(Calendar.MILLISECOND);
            //signedSignatureProperties.SigningTime = (/*setter*/xmlGregorianCalendar);

            //// SigningCertificate
            //if (signatureConfig.SigningCertificateChain == null
            //    || signatureConfig.SigningCertificateChain.IsEmpty())
            //{
            //    throw new Exception("no signing certificate chain available");
            //}
            //CertIDListType signingCertificates = signedSignatureProperties.AddNewSigningCertificate();
            //CertIDType certId = signingCertificates.AddNewCert();
            //X509Certificate certificate = signatureConfig.SigningCertificateChain.Get(0);
            //SetCertID(certId, signatureConfig, signatureConfig.IsXadesIssuerNameNoReverseOrder(), certificate);

            //// ClaimedRole
            //String role = signatureConfig.XadesRole;
            //if (role != null && !role.IsEmpty())
            //{
            //    SignerRoleType signerRole = signedSignatureProperties.AddNewSignerRole();
            //    signedSignatureProperties.SignerRole = (/*setter*/signerRole);
            //    ClaimedRolesListType claimedRolesList = signerRole.AddNewClaimedRoles();
            //    AnyType claimedRole = claimedRolesList.AddNewClaimedRole();
            //    XmlString roleString = XmlString.Factory.NewInstance();
            //    roleString.StringValue = (/*setter*/role);
            //    insertXChild(claimedRole, roleString);
            //}

            //// XAdES-EPES
            //SignaturePolicyService policyService = signatureConfig.SignaturePolicyService;
            //if (policyService != null)
            //{
            //    SignaturePolicyIdentifierType signaturePolicyIdentifier =
            //        signedSignatureProperties.AddNewSignaturePolicyIdentifier();

            //    SignaturePolicyIdType signaturePolicyId = signaturePolicyIdentifier.AddNewSignaturePolicyId();

            //    ObjectIdentifierType objectIdentifier = signaturePolicyId.AddNewSigPolicyId();
            //    objectIdentifier.Description = (/*setter*/policyService.SignaturePolicyDescription);

            //    IdentifierType identifier = objectIdentifier.AddNewIdentifier();
            //    identifier.StringValue = (/*setter*/policyService.SignaturePolicyIdentifier);

            //    byte[] signaturePolicyDocumentData = policyService.SignaturePolicyDocument;
            //    DigestAlgAndValueType sigPolicyHash = signaturePolicyId.AddNewSigPolicyHash();
            //    SetDigestAlgAndValue(sigPolicyHash, signaturePolicyDocumentData, signatureConfig.DigestAlgo);

            //    String signaturePolicyDownloadUrl = policyService.SignaturePolicyDownloadUrl;
            //    if (null != signaturePolicyDownloadUrl)
            //    {
            //        SigPolicyQualifiersListType sigPolicyQualifiers = signaturePolicyId.AddNewSigPolicyQualifiers();
            //        AnyType sigPolicyQualifier = sigPolicyQualifiers.AddNewSigPolicyQualifier();
            //        XmlString spUriElement = XmlString.Factory.NewInstance();
            //        spUriElement.StringValue = (/*setter*/signaturePolicyDownloadUrl);
            //        insertXChild(sigPolicyQualifier, spUriElement);
            //    }
            //}
            //else if (signatureConfig.IsXadesSignaturePolicyImplied())
            //{
            //    SignaturePolicyIdentifierType signaturePolicyIdentifier =
            //            signedSignatureProperties.AddNewSignaturePolicyIdentifier();
            //    signaturePolicyIdentifier.AddNewSignaturePolicyImplied();
            //}

            //// DataObjectFormat
            //if (!dataObjectFormatMimeTypes.IsEmpty())
            //{
            //    SignedDataObjectPropertiesType signedDataObjectProperties =
            //        signedProperties.AddNewSignedDataObjectProperties();

            //    List<DataObjectFormatType> dataObjectFormats = signedDataObjectProperties
            //            .DataObjectFormatList;
            //    foreach (Map.Entry<String, String> dataObjectFormatMimeType in this.dataObjectFormatMimeTypes
            //            .entrySet())
            //    {
            //        DataObjectFormatType dataObjectFormat = DataObjectFormatType.Factory.NewInstance();
            //        dataObjectFormat.ObjectReference = (/*setter*/"#" + dataObjectFormatMimeType.Key);
            //        dataObjectFormat.MimeType = (/*setter*/dataObjectFormatMimeType.Value);
            //        dataObjectFormats.Add(dataObjectFormat);
            //    }
            //}

            //// add XAdES ds:Object
            //List<XMLStructure> xadesObjectContent = new List<XMLStructure>();
            //Element qualDocElSrc = (Element)qualifyingProperties.DomNode;
            //Element qualDocEl = (Element)document.ImportNode(qualDocElSrc, true);
            //xadesObjectContent.Add(new DOMStructure(qualDocEl));
            //XMLObject xadesObject = GetSignatureFactory().newXMLObject(xadesObjectContent, null, null, null);
            //objects.Add(xadesObject);

            //// add XAdES ds:Reference
            //List<Transform> transforms = new List<Transform>();
            //Transform exclusiveTransform = newTransform(CanonicalizationMethod.INCLUSIVE);
            //transforms.Add(exclusiveTransform);
            //Reference reference = newReference
            //    ("#" + signatureConfig.XadesSignatureId, transforms, XADES_TYPE, null, null);
            //references.Add(reference);
            throw new NotImplementedException();
        }

        /**
         * Gives back the JAXB DigestAlgAndValue data structure.
         *
         * @param digestAlgAndValue the parent for the new digest element 
         * @param data the data to be digested
         * @param digestAlgo the digest algorithm
         */
        //protected static void SetDigestAlgAndValue(
        //        DigestAlgAndValueType digestAlgAndValue,
        //        byte[] data,
        //        HashAlgorithm digestAlgo)
        //{
        //    DigestMethodType digestMethod = digestAlgAndValue.AddNewDigestMethod();
        //    digestMethod.Algorithm = (/*setter*/SignatureConfig.GetDigestMethodUri(digestAlgo));

        //    MessageDigest messageDigest = CryptoFunctions.GetMessageDigest(digestAlgo);
        //    byte[] digestValue = messageDigest.Digest(data);
        //    digestAlgAndValue.DigestValue = (/*setter*/digestValue);
        //}

        /**
         * Gives back the JAXB CertID data structure.
         */
        //protected static void SetCertID
        //    (CertIDType certId, SignatureConfig signatureConfig, bool IssuerNameNoReverseOrder, X509Certificate certificate)
        //{
        //    X509IssuerSerialType issuerSerial = certId.AddNewIssuerSerial();
        //    String issuerName;
        //    if (issuerNameNoReverseOrder)
        //    {
        //        /*
        //         * Make sure the DN is encoded using the same order as present
        //         * within the certificate. This is an Office2010 work-around.
        //         * Should be reverted back.
        //         * 
        //         * XXX: not correct according to RFC 4514.
        //         */
        //        // TODO: check if issuerName is different on GetTBSCertificate
        //        // issuerName = PrincipalUtil.GetIssuerX509Principal(certificate).Name.Replace(",", ", ");
        //        issuerName = certificate.IssuerDN.Name.Replace(",", ", ");
        //    }
        //    else
        //    {
        //        issuerName = certificate.IssuerX500Principal.ToString();
        //    }
        //    issuerSerial.X509IssuerName = (/*setter*/issuerName);
        //    issuerSerial.X509SerialNumber = (/*setter*/certificate.SerialNumber);

        //    byte[] encodedCertificate;
        //    try
        //    {
        //        encodedCertificate = certificate.Encoded;
        //    }
        //    catch (CertificateEncodingException e)
        //    {
        //        throw new Exception("certificate encoding error: "
        //                + e.Message, e);
        //    }
        //    DigestAlgAndValueType certDigest = certId.AddNewCertDigest();
        //    SetDigestAlgAndValue(certDigest, encodedCertificate, signatureConfig.XadesDigestAlgo);
        //}

        /**
         * Adds a mime-type for the given ds:Reference (referred via its @URI). This
         * information is Added via the xades:DataObjectFormat element.
         * 
         * @param dsReferenceUri
         * @param mimetype
         */
        public void AddMimeType(String dsReferenceUri, String mimetype)
        {
            this.dataObjectFormatMimeTypes.Add(dsReferenceUri, mimetype);
        }

        protected static void insertXChild(XmlNode root, XmlNode child)
        {
            throw new NotImplementedException();
            //XmlCursor rootCursor = root.NewCursor();
            //rootCursor.ToEndToken();
            //XmlCursor childCursor = child.NewCursor();
            //childCursor.ToNextToken();
            //childCursor.MoveXml(rootCursor);
            //childCursor.Dispose();
            //rootCursor.Dispose();
        }

    }
}
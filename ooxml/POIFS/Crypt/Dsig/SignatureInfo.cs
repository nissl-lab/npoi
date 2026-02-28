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

namespace NPOI.POIFS.Crypt.Dsig
{
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Security.Cryptography.X509Certificates;
    using System.Xml;
    using System.Xml.XPath;

    /**
     * <p>This class is the default entry point for XML signatures and can be used for
     * validating an existing signed office document and signing a office document.</p>
     * 
     * <p><b>Validating a signed office document</b></p>
     * 
     * <pre>
     * OPCPackage pkg = OPCPackage.open(..., PackageAccess.READ);
     * SignatureConfig sic = new SignatureConfig();
     * sic.setOpcPackage(pkg);
     * SignatureInfo si = new SignatureInfo();
     * si.setSignatureConfig(sic);
     * boolean isValid = si.validate();
     * ...
     * </pre>
     * 
     * <p><b>Signing an office document</b></p>
     * 
     * <pre>
     * // loading the keystore - pkcs12 is used here, but of course jks &amp; co are also valid
     * // the keystore needs to contain a private key and it's certificate having a
     * // 'digitalSignature' key usage
     * char password[] = "test".toCharArray();
     * File file = new File("test.pfx");
     * KeyStore keystore = KeyStore.getInstance("PKCS12");
     * FileInputStream fis = new FileInputStream(file);
     * keystore.load(fis, password);
     * fis.close();
     * 
     * // extracting private key and certificate
     * String alias = "xyz"; // alias of the keystore entry
     * Key key = keystore.getKey(alias, password);
     * X509Certificate x509 = (X509Certificate)keystore.getCertificate(alias);
     * 
     * // filling the SignatureConfig entries (minimum fields, more options are available ...)
     * SignatureConfig signatureConfig = new SignatureConfig();
     * signatureConfig.setKey(keyPair.getPrivate());
     * signatureConfig.setSigningCertificateChain(Collections.singletonList(x509));
     * OPCPackage pkg = OPCPackage.open(..., PackageAccess.READ_WRITE);
     * signatureConfig.setOpcPackage(pkg);
     * 
     * // adding the signature document to the package
     * SignatureInfo si = new SignatureInfo();
     * si.setSignatureConfig(signatureConfig);
     * si.confirmSignature();
     * // optionally verify the generated signature
     * boolean b = si.verifySignature();
     * assert (b);
     * // write the changes back to disc
     * pkg.close();
     * </pre>
     * 
     * <p><b>Implementation notes:</b></p>
     * 
     * <p>Although there's a XML signature implementation in the Oracle JDKs 6 and higher,
     * compatibility with IBM JDKs is also in focus (... but maybe not thoroughly tested ...).
     * Therefore we are using the Apache Santuario libs (xmlsec) instead of the built-in classes,
     * as the compatibility seems to be provided there.</p>
     * 
     * <p>To use SignatureInfo and its sibling classes, you'll need to have the following libs
     * in the classpath:</p>
     * <ul>
     * <li>BouncyCastle bcpkix and bcprov (tested against 1.51)</li>
     * <li>Apache Santuario "xmlsec" (tested against 2.0.1)</li>
     * <li>and slf4j-api (tested against 1.7.7)</li>
     * </ul>
     */
    public class SignatureInfo : ISignatureConfigurable
    {

        //private static POILogger LOG = POILogFactory.GetLogger(typeof(SignatureInfo));
        private static bool IsInitialized = false;

        protected internal SignatureConfig signatureConfig;

        public class SignaturePart
        {
            private PackagePart signaturePart;
            private X509Certificate signer;
            private List<X509Certificate> certChain;

            private SignaturePart(PackagePart signaturePart)
            {
                this.signaturePart = signaturePart;
            }

            /**
             * @return the package part Containing the signature
             */
            public PackagePart GetPackagePart()
            {
                return signaturePart;
            }

            /**
             * @return the signer certificate
             */
            public X509Certificate GetSigner()
            {
                return signer;
            }

            /**
             * @return the certificate chain of the signer
             */
            public List<X509Certificate> GetCertChain()
            {
                return certChain;
            }

            /**
             * Helper method for examining the xml signature
             *
             * @return the xml signature document
             * @throws IOException if the xml signature doesn't exist or can't be read
             * @throws XmlException if the xml signature is malformed
             */
            //public SignatureDocument GetSignatureDocument() {
            //    // TODO: check for XXE
            //    return SignatureDocument.Parse(signaturePart.InputStream);
            //}

            /**
             * @return true, when the xml signature is valid, false otherwise
             * 
             * @throws EncryptedDocumentException if the signature can't be extracted or if its malformed
             */
            public bool Validate()
            {
                //    KeyInfoKeySelector keySelector = new KeyInfoKeySelector();
                //    try {
                //        XPathDocument doc = DocumentHelper.ReadDocument(signaturePart.InputStream);
                //        XPath xpath = XPathFactory.NewInstance().newXPath();
                //        NodeList nl = (NodeList)xpath.Compile("//*[@Id]").Evaluate(doc, XPathConstants.NODESET);
                //        for (int i = 0; i < nl.Length; i++) {
                //            ((Element)nl.Item(i)).IdAttribute = (/*setter*/"Id", true);
                //        }

                //        DOMValidateContext domValidateContext = new DOMValidateContext(keySelector, doc);
                //        domValidateContext.Property = (/*setter*/"org.jcp.xml.dsig.validateManifests", Boolean.TRUE);
                //        domValidateContext.URIDereferencer = (/*setter*/signatureConfig.UriDereferencer);
                //        brokenJvmWorkaround(domValidateContext);

                //        XMLSignatureFactory xmlSignatureFactory = signatureConfig.SignatureFactory;
                //        XMLSignature xmlSignature = xmlSignatureFactory.UnmarshalXMLSignature(domValidateContext);

                //        // TODO: replace with property when xml-sec patch is applied
                //        foreach (Reference ref1 in (List<Reference>)xmlSignature.SignedInfo.References) {
                //            SignatureFacet.BrokenJvmWorkaround(ref1);
                //        }
                //        foreach (XMLObject xo in (List<XMLObject>)xmlSignature.Objects) {
                //            foreach (XMLStructure xs in (List<XMLStructure>)xo.Content) {
                //                if (xs is Manifest) {
                //                    foreach (Reference ref1 in (List<Reference>)((Manifest)xs).References) {
                //                        SignatureFacet.BrokenJvmWorkaround(ref1);
                //                    }
                //                }
                //            }
                //        }

                //        bool valid = xmlSignature.Validate(domValidateContext);

                //        if (valid) {
                //            signer = keySelector.Signer;
                //            certChain = keySelector.CertChain;
                //        }

                //        return valid;
                //    } catch (Exception e) {
                //        String s = "error in marshalling and validating the signature";
                //        LOG.Log(POILogger.ERROR, s, e);
                //        throw new EncryptedDocumentException(s, e);
                //    }
                throw new NotImplementedException();
            }

        }

        /**
         * Constructor Initializes xml signature environment, if it hasn't been Initialized before
         */
        public SignatureInfo()
        {
            InitXmlProvider();
        }

        /**
         * @return the signature config
         */
        public SignatureConfig GetSignatureConfig()
        {
            //return signatureConfig;
            throw new NotImplementedException();
        }

        /**
         * @param signatureConfig the signature config, needs to be Set before a SignatureInfo object is used
         */
        public void SetSignatureConfig(SignatureConfig signatureConfig)
        {
            //this.signatureConfig = signatureConfig;
            throw new NotImplementedException();
        }

        /**
         * @return true, if first signature part is valid
         */
        public bool VerifySignature()
        {
            // http://www.oracle.com/technetwork/articles/javase/dig-signature-api-140772.html
            //foreach (SignaturePart sp in GetSignatureParts()) {
            //    // only validate first part
            //    return sp.Validate();
            //}
            //return false;
            throw new NotImplementedException();
        }

        /**
         * add the xml signature to the document
         *
         * @throws XMLSignatureException
         * @throws MarshalException
         */
        public void ConfirmSignature()
        {
            XPathDocument document = DocumentHelper.CreateDocument();

            // operate
            //DigestInfo digestInfo = preSign(document, null);

            // Setup: key material, signature value
            //byte[] signatureValue = signDigest(digestInfo.digestValue);

            // operate: postSign
            //postSign(document, signatureValue);
            throw new NotImplementedException();
        }

        /**
         * Sign (encrypt) the digest with the private key.
         * Currently only rsa is supported.
         *
         * @param digest the hashed input
         * @return the encrypted hash
         */
        public byte[] signDigest(byte[] digest)
        {
            Cipher cipher = CryptoFunctions.GetCipher(signatureConfig.GetKey(), CipherAlgorithm.rsa
                , ChainingMode.ecb, null, Cipher.ENCRYPT_MODE, "PKCS1PAdding");

            //try {
            //    MemoryStream digestInfoValueBuf = new MemoryStream();
            //    digestInfoValueBuf.Write(signatureConfig.GetHashMagic());
            //    digestInfoValueBuf.Write(digest);
            //    byte[] digestInfoValue = digestInfoValueBuf.ToArray();
            //    byte[] signatureValue = cipher.DoFinal(digestInfoValue);
            //    return signatureValue;
            //} catch (Exception e) {
            //    throw new EncryptedDocumentException(e);
            //}
            throw new NotImplementedException();
        }

        /**
         * @return a signature part for each signature document.
         * the parts can be validated independently.
         */
        public IEnumerable<SignaturePart> GetSignatureParts()
        {
            signatureConfig.Init(true);
            throw new NotImplementedException();
            //return new Iterable<SignaturePart>() {
            //    public Iterator<SignaturePart> iterator() {
            //        return new Iterator<SignaturePart>() {
            //            OPCPackage pkg = signatureConfig.OpcPackage;
            //            Iterator<PackageRelationship> sigOrigRels = 
            //                pkg.GetRelationshipsByType(PackageRelationshipTypes.DIGITAL_SIGNATURE_ORIGIN).iterator();
            //            Iterator<PackageRelationship> sigRels = null;
            //            PackagePart sigPart = null;

            //            public bool HasNext() {
            //                while (sigRels == null || !sigRels.HasNext()) {
            //                    if (!sigOrigRels.HasNext()) return false;
            //                    sigPart = pkg.GetPart(sigOrigRels.Next());
            //                    LOG.Log(POILogger.DEBUG, "Digital Signature Origin part", sigPart);
            //                    try {
            //                        sigRels = sigPart.GetRelationshipsByType(PackageRelationshipTypes.DIGITAL_SIGNATURE).iterator();
            //                    } catch (InvalidFormatException e) {
            //                        LOG.Log(POILogger.WARN, "Reference to signature is invalid.", e);
            //                    }
            //                }
            //                return true;
            //            }

            //            public SignaturePart next() {
            //                PackagePart sigRelPart = null;
            //                do {
            //                    try {
            //                        if (!hasNext()) throw new NoSuchElementException();
            //                        sigRelPart = sigPart.GetRelatedPart(sigRels.Next()); 
            //                        LOG.Log(POILogger.DEBUG, "XML Signature part", sigRelPart);
            //                    } catch (InvalidFormatException e) {
            //                        LOG.Log(POILogger.WARN, "Reference to signature is invalid.", e);
            //                    }
            //                } while (sigPart == null);
            //                return new SignaturePart(sigRelPart);
            //            }

            //            public void Remove() {
            //                throw new UnsupportedOperationException();
            //            }
            //        };
            //    }
            //};
        }

        /**
         * Initialize the xml signing environment and the bouncycastle provider 
         */
        protected static void InitXmlProvider()
        {
            //if (isInitialized) return;
            //isInitialized = true;

            //try {
            //    Init.Init();
            //    RelationshipTransformService.RegisterDsigProvider();
            //    CryptoFunctions.RegisterBouncyCastle();
            //} catch (Exception e) {
            //    throw new Exception("Xml & BouncyCastle-Provider Initialization failed", e);
            //}
            throw new NotImplementedException();
        }

        /**
         * Helper method for Adding informations before the signing.
         * Normally {@link #ConfirmSignature()} is sufficient to be used.
         */
        public DigestInfo preSign(XmlDocument document, List<DigestInfo> digestInfos)
        {
            signatureConfig.Init(false);

            //// it's necessary to explicitly Set the mdssi namespace, but the sign() method has no
            //// normal way to interfere with, so we need to add the namespace under the hand ...
            //EventTarget target = (EventTarget)document;
            //EventListener creationListener = signatureConfig.SignatureMarshalListener;
            //if (creationListener != null) {
            //    if (creationListener is SignatureMarshalListener) {
            //        ((SignatureMarshalListener)creationListener).EventTarget = (/*setter*/target);
            //    }
            //    SignatureMarshalListener.Listener = (/*setter*/target, creationListener, true);
            //}

            ///*
            // * Signature context construction.
            // */
            //XMLSignContext xmlSignContext = new DOMSignContext(signatureConfig.Key, document);
            //URIDereferencer uriDereferencer = signatureConfig.UriDereferencer;
            //if (null != uriDereferencer) {
            //    xmlSignContext.URIDereferencer = (/*setter*/uriDereferencer);
            //}

            //foreach (Map.Entry<String, String> me in signatureConfig.NamespacePrefixes.EntrySet()) {
            //    xmlSignContext.PutNamespacePrefix(me.Key, me.Value);
            //}
            //xmlSignContext.DefaultNamespacePrefix = (/*setter*/"");
            //// signatureConfig.NamespacePrefixes.Get(XML_DIGSIG_NS));

            //brokenJvmWorkaround(xmlSignContext);

            //XMLSignatureFactory signatureFactory = signatureConfig.SignatureFactory;

            ///*
            // * Add ds:References that come from signing client local files.
            // */
            //List<Reference> references = new List<Reference>();
            //foreach (DigestInfo digestInfo in safe(digestInfos)) {
            //    byte[] documentDigestValue = digestInfo.digestValue;

            //    String uri = new File(digestInfo.description).Name;
            //    Reference reference = SignatureFacet.newReference
            //        (uri, null, null, null, documentDigestValue, signatureConfig);
            //    references.Add(reference);
            //}

            ///*
            // * Invoke the signature facets.
            // */
            //List<XMLObject> objects = new List<XMLObject>();
            //foreach (SignatureFacet signatureFacet in signatureConfig.SignatureFacets) {
            //    LOG.Log(POILogger.DEBUG, "invoking signature facet: " + signatureFacet.Class.SimpleName);
            //    signatureFacet.PreSign(document, references, objects);
            //}

            ///*
            // * ds:SignedInfo
            // */
            //SignedInfo signedInfo;
            //try {
            //    SignatureMethod signatureMethod = signatureFactory.newSignatureMethod
            //        (signatureConfig.SignatureMethodUri, null);
            //    CanonicalizationMethod canonicalizationMethod = signatureFactory
            //        .newCanonicalizationMethod(signatureConfig.CanonicalizationMethod,
            //        (C14NMethodParameterSpec)null);
            //    signedInfo = signatureFactory.NewSignedInfo(
            //        canonicalizationMethod, signatureMethod, references);
            //} catch (GeneralSecurityException e) {
            //    throw new XMLSignatureException(e);
            //}

            ///*
            // * JSR105 ds:Signature creation
            // */
            //String signatureValueId = signatureConfig.PackageSignatureId + "-signature-value";
            //javax.xml.Crypto.dsig.XMLSignature xmlSignature = signatureFactory
            //    .newXMLSignature(signedInfo, null, objects, signatureConfig.PackageSignatureId,
            //    signatureValueId);

            ///*
            // * ds:Signature Marshalling.
            // */
            //xmlSignature.Sign(xmlSignContext);

            ///*
            // * Completion of undigested ds:References in the ds:Manifests.
            // */
            //foreach (XMLObject object1 in objects) {
            //    LOG.Log(POILogger.DEBUG, "object java type: " + object1.Class.Name);
            //    List<XMLStructure> objectContentList = object1.Content;
            //    foreach (XMLStructure objectContent in objectContentList) {
            //        LOG.Log(POILogger.DEBUG, "object content java type: " + objectContent.Class.Name);
            //        if (!(objectContent is Manifest)) continue;
            //        Manifest manifest = (Manifest)objectContent;
            //        List<Reference> manifestReferences = manifest.References;
            //        foreach (Reference manifestReference in manifestReferences) {
            //            if (manifestReference.DigestValue != null) continue;

            //            DOMReference manifestDOMReference = (DOMReference)manifestReference;
            //            manifestDOMReference.Digest(xmlSignContext);
            //        }
            //    }
            //}

            ///*
            // * Completion of undigested ds:References.
            // */
            //List<Reference> signedInfoReferences = signedInfo.References;
            //foreach (Reference signedInfoReference in signedInfoReferences) {
            //    DOMReference domReference = (DOMReference)signedInfoReference;

            //    // ds:Reference with external digest value
            //    if (domReference.DigestValue != null) continue;

            //    domReference.Digest(xmlSignContext);
            //}

            ///*
            // * Calculation of XML signature digest value.
            // */
            //DOMSignedInfo domSignedInfo = (DOMSignedInfo)signedInfo;
            //MemoryStream dataStream = new MemoryStream();
            //domSignedInfo.Canonicalize(xmlSignContext, dataStream);
            //byte[] octets = dataStream.ToByteArray();

            ///*
            // * TODO: we could be using DigestOutputStream here to optimize memory
            // * usage.
            // */

            //MessageDigest md = CryptoFunctions.GetMessageDigest(signatureConfig.DigestAlgo);
            //byte[] digestValue = md.Digest(octets);


            //String description = signatureConfig.SignatureDescription;
            //return new DigestInfo(digestValue, signatureConfig.DigestAlgo, description);
            throw new NotImplementedException();
        }

        /**
         * Helper method for Adding informations After the signing.
         * Normally {@link #ConfirmSignature()} is sufficient to be used.
         */
        public void postSign(XmlDocument document, byte[] signatureValue)
        {
            //LOG.Log(POILogger.DEBUG, "postSign");

            /*
             * Check ds:Signature node.
             */
            //String signatureId = signatureConfig.PackageSignatureId;
            //if (!signatureId.Equals(document.DocumentElement.GetAttribute("Id"))) {
            //    throw new Exception("ds:Signature not found for @Id: " + signatureId);
            //}

            ///*
            // * Insert signature value into the ds:SignatureValue element
            // */
            //NodeList sigValNl = document.GetElementsByTagNameNS(XML_DIGSIG_NS, "SignatureValue");
            //if (sigValNl.Length != 1) {
            //    throw new Exception("preSign has to be called before postSign");
            //}
            //sigValNl.Item(0).TextContent = (/*setter*/Base64.Encode(signatureValue));

            ///*
            // * Allow signature facets to inject their own stuff.
            // */
            //foreach (SignatureFacet signatureFacet in signatureConfig.SignatureFacets) {
            //    signatureFacet.PostSign(document);
            //}

            //WriteDocument(document);
            throw new NotImplementedException();
        }

        /**
         * Write XML signature into the OPC package
         *
         * @param document the xml signature document
         * @throws MarshalException
         */
        protected void WriteDocument(XmlDocument document)
        {
            //XmlOptions xo = new XmlOptions();
            //Dictionary<String, String> namespaceMap = new Dictionary<String, String>();
            //foreach (Map.Entry<String, String> entry in signatureConfig.NamespacePrefixes.EntrySet()) {
            //    namespaceMap.Put(entry.Value, entry.Key);
            //}
            //xo.SaveSuggestedPrefixes = (/*setter*/namespaceMap);
            //xo.UseDefaultNamespace();

            //LOG.Log(POILogger.DEBUG, "output signed Office OpenXML document");

            ///*
            // * Copy the original OOXML content to the signed OOXML package. During
            // * copying some files need to Changed.
            // */
            //OPCPackage pkg = signatureConfig.OpcPackage;

            //PackagePartName sigPartName, sigsPartName;
            //try {
            //    // <Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.Openxmlformats-package.digital-signature-xmlsignature+xml"/>
            //    sigPartName = PackagingURIHelper.CreatePartName("/_xmlsignatures/sig1.xml");
            //    // <Default Extension="sigs" ContentType="application/vnd.Openxmlformats-package.digital-signature-origin"/>
            //    sigsPartName = PackagingURIHelper.CreatePartName("/_xmlsignatures/origin.sigs");
            //} catch (InvalidFormatException e) {
            //    throw new MarshalException(e);
            //}

            //PackagePart sigPart = pkg.GetPart(sigPartName);
            //if (sigPart == null) {
            //    sigPart = pkg.CreatePart(sigPartName, ContentTypes.DIGITAL_SIGNATURE_XML_SIGNATURE_PART);
            //}

            //try {
            //    OutputStream os = sigPart.OutputStream;
            //    SignatureDocument sigDoc = SignatureDocument.Factory.Parse(document);
            //    sigDoc.Save(os, xo);
            //    os.Close();
            //} catch (Exception e) {
            //    throw new MarshalException("Unable to write signature document", e);
            //}

            //PackagePart sigsPart = pkg.GetPart(sigsPartName);
            //if (sigsPart == null) {
            //    // touch empty marker file
            //    sigsPart = pkg.CreatePart(sigsPartName, ContentTypes.DIGITAL_SIGNATURE_ORIGIN_PART);
            //}

            //PackageRelationshipCollection relCol = pkg.GetRelationshipsByType(PackageRelationshipTypes.DIGITAL_SIGNATURE_ORIGIN);
            //foreach (PackageRelationship pr in relCol) {
            //    pkg.RemoveRelationship(pr.Id);
            //}
            //pkg.AddRelationship(sigsPartName, TargetMode.INTERNAL, PackageRelationshipTypes.DIGITAL_SIGNATURE_ORIGIN);

            //sigsPart.AddRelationship(sigPartName, TargetMode.INTERNAL, PackageRelationshipTypes.DIGITAL_SIGNATURE);
        }

        /**
         * Helper method for null lists, which are Converted to empty lists
         *
         * @param other the reference to wrap, if null
         * @return if other is null, an empty lists is returned, otherwise other is returned
         */
        private static List<T> safe<T>(List<T> other)
        {
            //return other == null ? Collections.EMPTY_LIST : other;
            throw new NotImplementedException();
        }

        //private void brokenJvmWorkaround(XMLSignContext context) {
        //    // workaround for https://bugzilla.redhat.com/Show_bug.cgi?id=1155012
        //    Provider bcProv = Security.GetProvider("BC");
        //    if (bcProv != null) {
        //        context.Property = (/*setter*/"org.jcp.xml.dsig.internal.dom.SignatureProvider", bcProv);
        //    }
        //}

        //private void brokenJvmWorkaround(XMLValidateContext context) {
        //    // workaround for https://bugzilla.redhat.com/Show_bug.cgi?id=1155012
        //    Provider bcProv = Security.GetProvider("BC");
        //    if (bcProv != null) {
        //        context.Property = (/*setter*/"org.jcp.xml.dsig.internal.dom.SignatureProvider", bcProv);
        //    }
        //}
    }

}
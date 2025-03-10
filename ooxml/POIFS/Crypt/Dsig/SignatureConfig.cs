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

namespace NPOI.POIFS.Crypt.Dsig
{
    using System;
    using System.Collections.Generic;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Cryptography.Xml;
    using System.Threading;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.Dsig.Facets;
    using NPOI.POIFS.Crypt.Dsig.Services;


    /**
     * This class bundles the configuration options used for the existing
     * signature facets.
     * Apart of the thread local members (e.g. opc-package) most values will probably be constant, so
     * it might be configured centrally (e.g. by spring) 
     */
    public class SignatureConfig
    {

        private ThreadLocal<OPCPackage> opcPackage = new ThreadLocal<OPCPackage>();
        //private ThreadLocal<XMLSignatureFactory> signatureFactory = new ThreadLocal<XMLSignatureFactory>();
        //private ThreadLocal<KeyInfoFactory> keyInfoFactory = new ThreadLocal<KeyInfoFactory>();
        //private ThreadLocal<Provider> provider = new ThreadLocal<Provider>();

        private List<SignatureFacet> signatureFacets = new List<SignatureFacet>();
        private HashAlgorithm digestAlgo = HashAlgorithm.sha1;
        private DateTime executionTime = DateTime.Now;
        private IPrivateKey key;
        private List<X509Certificate> signingCertificateChain;

        /**
         * the optional signature policy service used for XAdES-EPES.
         */
        private ISignaturePolicyService signaturePolicyService;
        private IURIDereferencer uriDereferencer = null;
        private String canonicalizationMethod = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";// CanonicalizationMethod.INCLUSIVE;

        private bool includeEntireCertificateChain = true;
        private bool includeIssuerSerial = false;
        private bool includeKeyValue = false;

        /**
         * the time-stamp service used for XAdES-T and XAdES-X.
         */
        private ITimeStampService tspService = new TSPTimeStampService();
        /**
         * timestamp service provider URL
         */
        private String tspUrl;
        private bool tspOldProtocol = false;
        /**
         * if not defined, it's the same as the main digest
         */
        private HashAlgorithm tspDigestAlgo = null;
        private String tspUser;
        private String tspPass;
        private ITimeStampServiceValidator tspValidator;
        /**
         * the optional TSP request policy OID.
         */
        private String tspRequestPolicy = "1.3.6.1.4.1.13762.3";
        private String userAgent = "POI XmlSign Service TSP Client";
        private String proxyUrl;

        /**
         * the optional revocation data service used for XAdES-C and XAdES-X-L.
         * When <code>null</code> the signature will be limited to XAdES-T only.
         */
        private IRevocationDataService revocationDataService;
        /**
         * if not defined, it's the same as the main digest
         */
        private HashAlgorithm xadesDigestAlgo = null;
        private String xadesRole = null;
        private String xadesSignatureId = "idSignedProperties";
        private bool xadesSignaturePolicyImplied = true;
        private String xadesCanonicalizationMethod = "http://www.w3.org/2001/10/xml-exc-c14n#";// CanonicalizationMethod.EXCLUSIVE;

        /**
         * Work-around for Office 2010 IssuerName encoding.
         */
        private bool xadesIssuerNameNoReverseOrder = true;

        /**
         * The signature Id attribute value used to create the XML signature. A
         * <code>null</code> value will trigger an automatically generated signature Id.
         */
        private String packageSignatureId = "idPackageSignature";

        /**
         * Gives back the human-readable description of what the citizen will be
         * signing. The default value is "Office OpenXML Document".
         */
        private String signatureDescription = "Office OpenXML Document";

        /**
         * The process of signing includes the marshalling of xml structures.
         * This also includes the canonicalization. Currently this leads to problems 
         * with certain namespaces, so this EventListener is used to interfere
         * with the marshalling Process.
         */
        IEventListener signatureMarshalListener = null;

        /**
         * Map of namespace uris to prefix
         * If a mapping is specified, the corresponding elements will be prefixed
         */
        Dictionary<String, String> namespacePrefixes = new Dictionary<String, String>();

        /**
         * Inits and Checks the config object.
         * If not Set previously, complex configuration properties also Get 
         * Created/initialized via this Initialization call.
         *
         * @param onlyValidation if true, only a subset of the properties
         * is Initialized, which are necessary for validation. If false,
         * also the other properties needed for signing are been taken care of
         */
        protected internal void Init(bool onlyValidation)
        {
            if (opcPackage == null)
            {
                throw new EncryptedDocumentException("opcPackage is null");
            }
            if (uriDereferencer == null)
            {
                uriDereferencer = new OOXMLURIDereferencer();
            }
            if (uriDereferencer is ISignatureConfigurable configurable)
            {
                configurable.SetSignatureConfig(this);
            }
            if (namespacePrefixes.Count == 0)
            {
                /*
                 * OOo doesn't like ds namespaces so per default prefixing is off.
                 */
                // namespacePrefixes.Put(XML_DIGSIG_NS, "");
                namespacePrefixes.Add(SignatureFacet.OO_DIGSIG_NS, "mdssi");
                namespacePrefixes.Add(SignatureFacet.XADES_132_NS, "xd");
            }

            if (onlyValidation) return;

            if (signatureMarshalListener == null)
            {
                signatureMarshalListener = new SignatureMarshalListener();
            }

            if (signatureMarshalListener is ISignatureConfigurable signatureConfigurable)
            {
                signatureConfigurable.SetSignatureConfig(this);
            }

            if (tspService != null)
            {
                tspService.SetSignatureConfig(this);
            }

            if (signatureFacets.Count==0)
            {
                AddSignatureFacet(new OOXMLSignatureFacet());
                AddSignatureFacet(new KeyInfoSignatureFacet());
                AddSignatureFacet(new XAdESSignatureFacet());
                AddSignatureFacet(new Office2010SignatureFacet());
            }

            foreach (SignatureFacet sf in signatureFacets)
            {
                sf.SetSignatureConfig(this);
            }
        }

        /**
         * @param signatureFacet the signature facet is Appended to facet list 
         */
        public void AddSignatureFacet(SignatureFacet signatureFacet)
        {
            signatureFacets.Add(signatureFacet);
        }

        /**
         * @return the list of facets, may be empty when the config object is not Initialized
         */
        public List<SignatureFacet> GetSignatureFacets()
        {
            return signatureFacets;
        }

        /**
         * @param signatureFacets the new list of facets
         */
        public void SetSignatureFacets(List<SignatureFacet> signatureFacets)
        {
            this.signatureFacets = signatureFacets;
        }

        /**
         * @return the main digest algorithm, defaults to sha-1
         */
        public HashAlgorithm GetDigestAlgo()
        {
            return digestAlgo;
        }

        /**
         * @param digestAlgo the main digest algorithm
         */
        public void SetDigestAlgo(HashAlgorithm digestAlgo)
        {
            this.digestAlgo = digestAlgo;
        }

        /**
         * @return the opc package to be used by this thread, stored as thread-local
         */
        public OPCPackage GetOpcPackage()
        {
            return opcPackage.Value;
        }

        /**
         * @param opcPackage the opc package to be handled by this thread, stored as thread-local
         */
        public void SetOpcPackage(OPCPackage opcPackage)
        {
            this.opcPackage.Value = opcPackage;
        }

        /**
         * @return the private key
         */
        public IPrivateKey GetKey()
        {
            return key;
        }

        /**
         * @param key the private key
         */
        public void SetKey(IPrivateKey key)
        {
            this.key = key;
        }

        /**
         * @return the certificate chain, index 0 is usually the certificate matching
         * the private key
         */
        public List<X509Certificate> GetSigningCertificateChain()
        {
            return signingCertificateChain;
        }

        /**
         * @param signingCertificateChain the certificate chain, index 0 should be
         * the certificate matching the private key
         */
        public void SetSigningCertificateChain(
                List<X509Certificate> signingCertificateChain)
        {
            this.signingCertificateChain = signingCertificateChain;
        }

        /**
         * @return the time at which the document is signed, also used for the timestamp service.
         * defaults to now
         */
        public DateTime GetExecutionTime()
        {
            return executionTime;
        }

        /**
         * @param executionTime Sets the time at which the document ought to be signed
         */
        public void SetExecutionTime(DateTime executionTime)
        {
            this.executionTime = executionTime;
        }

        /**
         * @return the service to be used for XAdES-EPES properties. There's no default implementation
         */
        public ISignaturePolicyService GetSignaturePolicyService()
        {
            return signaturePolicyService;
        }

        /**
         * @param signaturePolicyService the service to be used for XAdES-EPES properties
         */
        public void SetSignaturePolicyService(ISignaturePolicyService signaturePolicyService)
        {
            this.signaturePolicyService = signaturePolicyService;
        }

        /**
         * @return the dereferencer used for Reference/@URI attributes, defaults to {@link OOXMLURIDereferencer}
         */
        //public URIDereferencer GetUriDereferencer()
        //{
        //    return uriDereferencer;
        //}

        /**
         * @param uriDereferencer the dereferencer used for Reference/@URI attributes
         */
        //public void SetUriDereferencer(IURIDereferencer uriDereferencer)
        //{
        //    this.uriDereferencer = uriDereferencer;
        //}

        /**
         * @return Gives back the human-readable description of what the citizen
         * will be signing. The default value is "Office OpenXML Document".
         */
        public String GetSignatureDescription()
        {
            return signatureDescription;
        }

        /**
         * @param signatureDescription the human-readable description of
         * what the citizen will be signing.
         */
        public void SetSignatureDescription(String signatureDescription)
        {
            this.signatureDescription = signatureDescription;
        }

        /**
         * @return the default canonicalization method, defaults to INCLUSIVE
         */
        public String GetCanonicalizationMethod()
        {
            return canonicalizationMethod;
        }

        /**
         * @param canonicalizationMethod the default canonicalization method
         */
        public void SetCanonicalizationMethod(String canonicalizationMethod)
        {
            this.canonicalizationMethod = canonicalizationMethod;
        }

        /**
         * @return The signature Id attribute value used to create the XML signature.
         * Defaults to "idPackageSignature"
         */
        public String GetPackageSignatureId()
        {
            return packageSignatureId;
        }

        /**
         * @param packageSignatureId The signature Id attribute value used to create the XML signature.
         * A <code>null</code> value will trigger an automatically generated signature Id.
         */
        public void SetPackageSignatureId(String packageSignatureId)
        {
            this.packageSignatureId = nvl(packageSignatureId, "xmldsig-" + Guid.NewGuid().ToString());// UUID.RandomUUID());
        }

        /**
         * @return the url of the timestamp provider (TSP)
         */
        public String GetTspUrl()
        {
            return tspUrl;
        }

        /**
         * @param tspUrl the url of the timestamp provider (TSP)
         */
        public void SetTspUrl(String tspUrl)
        {
            this.tspUrl = tspUrl;
        }

        /**
         * @return if true, uses timestamp-request/response mimetype,
         * if false, timestamp-query/reply mimetype 
         */
        public bool IsTspOldProtocol()
        {
            return tspOldProtocol;
        }

        /**
         * @param tspOldProtocol defines the timestamp-protocol mimetype
         * @see #isTspOldProtocol
         */
        public void SetTspOldProtocol(bool tspOldProtocol)
        {
            this.tspOldProtocol = tspOldProtocol;
        }

        /**
         * @return the hash algorithm to be used for the timestamp entry.
         * Defaults to the hash algorithm of the main entry
         */
        public HashAlgorithm GetTspDigestAlgo()
        {
            return nvl(tspDigestAlgo, digestAlgo);
        }

        /**
         * @param tspDigestAlgo the algorithm to be used for the timestamp entry.
         * if <code>null</code>, the hash algorithm of the main entry
         */
        public void SetTspDigestAlgo(HashAlgorithm tspDigestAlgo)
        {
            this.tspDigestAlgo = tspDigestAlgo;
        }

        /**
         * @return the proxy url to be used for all communications.
         * Currently this affects the timestamp service
         */
        public String GetProxyUrl()
        {
            return proxyUrl;
        }

        /**
         * @param proxyUrl the proxy url to be used for all communications.
         * Currently this affects the timestamp service
         */
        public void SetProxyUrl(String proxyUrl)
        {
            this.proxyUrl = proxyUrl;
        }

        /**
         * @return the timestamp service. Defaults to {@link TSPTimeStampService}
         */
        public ITimeStampService GetTspService()
        {
            return tspService;
        }

        /**
         * @param tspService the timestamp service
         */
        public void SetTspService(ITimeStampService tspService)
        {
            this.tspService = tspService;
        }

        /**
         * @return the user id for the timestamp service - currently only basic authorization is supported
         */
        public String GetTspUser()
        {
            return tspUser;
        }

        /**
         * @param tspUser the user id for the timestamp service - currently only basic authorization is supported
         */
        public void SetTspUser(String tspUser)
        {
            this.tspUser = tspUser;
        }

        /**
         * @return the password for the timestamp service
         */
        public String GetTspPass()
        {
            return tspPass;
        }

        /**
         * @param tspPass the password for the timestamp service
         */
        public void SetTspPass(String tspPass)
        {
            this.tspPass = tspPass;
        }

        /**
         * @return the validator for the timestamp service (certificate)
         */
        public ITimeStampServiceValidator GetTspValidator()
        {
            return tspValidator;
        }

        /**
         * @param tspValidator the validator for the timestamp service (certificate)
         */
        public void SetTspValidator(ITimeStampServiceValidator tspValidator)
        {
            this.tspValidator = tspValidator;
        }

        /**
         * @return the optional revocation data service used for XAdES-C and XAdES-X-L.
         * When <code>null</code> the signature will be limited to XAdES-T only.
         */
        public IRevocationDataService GetRevocationDataService()
        {
            return revocationDataService;
        }

        /**
         * @param revocationDataService the optional revocation data service used for XAdES-C and XAdES-X-L.
         * When <code>null</code> the signature will be limited to XAdES-T only.
         */
        public void SetRevocationDataService(IRevocationDataService revocationDataService)
        {
            this.revocationDataService = revocationDataService;
        }

        /**
         * @return hash algorithm used for XAdES. Defaults to the {@link #getDigestAlgo()}
         */
        public HashAlgorithm GetXadesDigestAlgo()
        {
            return nvl(xadesDigestAlgo, digestAlgo);
        }

        /**
         * @param xadesDigestAlgo hash algorithm used for XAdES.
         * When <code>null</code>, defaults to {@link #getDigestAlgo()}
         */
        public void SetXadesDigestAlgo(HashAlgorithm xadesDigestAlgo)
        {
            this.xadesDigestAlgo = xadesDigestAlgo;
        }

        /**
         * @return the user agent used for http communication (e.g. to the TSP)
         */
        public String GetUserAgent()
        {
            return userAgent;
        }

        /**
         * @param userAgent the user agent used for http communication (e.g. to the TSP)
         */
        public void SetUserAgent(String userAgent)
        {
            this.userAgent = userAgent;
        }

        /**
         * @return the asn.1 object id for the tsp request policy.
         * Defaults to <code>1.3.6.1.4.1.13762.3</code>
         */
        public String GetTspRequestPolicy()
        {
            return tspRequestPolicy;
        }

        /**
         * @param tspRequestPolicy the asn.1 object id for the tsp request policy.
         */
        public void SetTspRequestPolicy(String tspRequestPolicy)
        {
            this.tspRequestPolicy = tspRequestPolicy;
        }

        /**
         * @return true, if the whole certificate chain is included in the signature.
         * When false, only the signer cert will be included 
         */
        public bool IsIncludeEntireCertificateChain()
        {
            return includeEntireCertificateChain;
        }

        /**
         * @param includeEntireCertificateChain if true, include the whole certificate chain.
         * If false, only include the signer cert
         */
        public void SetIncludeEntireCertificateChain(bool includeEntireCertificateChain)
        {
            this.includeEntireCertificateChain = includeEntireCertificateChain;
        }

        /**
         * @return if true, issuer serial number is included
         */
        public bool IsIncludeIssuerSerial()
        {
            return includeIssuerSerial;
        }

        /**
         * @param includeIssuerSerial if true, issuer serial number is included
         */
        public void SetIncludeIssuerSerial(bool includeIssuerSerial)
        {
            this.includeIssuerSerial = includeIssuerSerial;
        }

        /**
         * @return if true, the key value of the public key (certificate) is included
         */
        public bool IsIncludeKeyValue()
        {
            return includeKeyValue;
        }

        /**
         * @param includeKeyValue if true, the key value of the public key (certificate) is included
         */
        public void SetIncludeKeyValue(bool includeKeyValue)
        {
            this.includeKeyValue = includeKeyValue;
        }

        /**
         * @return the xades role element. If <code>null</code> the claimed role element is omitted.
         * Defaults to <code>null</code>
         */
        public String GetXadesRole()
        {
            return xadesRole;
        }

        /**
         * @param xadesRole the xades role element. If <code>null</code> the claimed role element is omitted.
         */
        public void SetXadesRole(String xadesRole)
        {
            this.xadesRole = xadesRole;
        }

        /**
         * @return the Id for the XAdES SignedProperties element.
         * Defaults to <code>idSignedProperties</code>
         */
        public String GetXadesSignatureId()
        {
            return nvl(xadesSignatureId, "idSignedProperties");
        }

        /**
         * @param xadesSignatureId the Id for the XAdES SignedProperties element.
         * When <code>null</code> defaults to <code>idSignedProperties</code>
         */
        public void SetXadesSignatureId(String xadesSignatureId)
        {
            this.xadesSignatureId = xadesSignatureId;
        }

        /**
         * @return when true, include the policy-implied block.
         * Defaults to <code>true</code>
         */
        public bool IsXadesSignaturePolicyImplied()
        {
            return xadesSignaturePolicyImplied;
        }

        /**
         * @param xadesSignaturePolicyImplied when true, include the policy-implied block
         */
        public void SetXadesSignaturePolicyImplied(bool xadesSignaturePolicyImplied)
        {
            this.xadesSignaturePolicyImplied = xadesSignaturePolicyImplied;
        }

        /**
         * Make sure the DN is encoded using the same order as present
         * within the certificate. This is an Office2010 work-around.
         * Should be reverted back.
         * 
         * XXX: not correct according to RFC 4514.
         *
         * @return when true, the issuer DN is used instead of the issuer X500 principal
         */
        public bool IsXadesIssuerNameNoReverseOrder()
        {
            return xadesIssuerNameNoReverseOrder;
        }

        /**
         * @param xadesIssuerNameNoReverseOrder when true, the issuer DN instead of the issuer X500 prinicpal is used
         */
        public void SetXadesIssuerNameNoReverseOrder(bool xadesIssuerNameNoReverseOrder)
        {
            this.xadesIssuerNameNoReverseOrder = xadesIssuerNameNoReverseOrder;
        }


        /**
         * @return the event listener which is active while xml structure for
         * the signature is Created.
         * Defaults to {@link SignatureMarshalListener}
         */
        //public EventListener GetSignatureMarshalListener()
        //{
        //    return signatureMarshalListener;
        //}

        /**
         * @param signatureMarshalListener the event listener watching the xml structure
         * generation for the signature
         */
        //public void SetSignatureMarshalListener(EventListener signatureMarshalListener)
        //{
        //    this.signatureMarshalListener = signatureMarshalListener;
        //}

        /**
         * @return the map of namespace uri (key) to prefix (value)
         */
        public Dictionary<String, String> GetNamespacePrefixes()
        {
            return namespacePrefixes;
        }

        /**
         * @param namespacePrefixes the map of namespace uri (key) to prefix (value)
         */
        public void SetNamespacePrefixes(Dictionary<String, String> namespacePrefixes)
        {
            this.namespacePrefixes = namespacePrefixes;
        }

        /**
         * helper method for null/default value handling
         * @param value
         * @param defaultValue
         * @return if value is not null, return value otherwise defaultValue
         */
        protected static T nvl<T>(T value, T defaultValue)
        {
            return value == null ? defaultValue : value;
        }

        /**
         * Each digest method has its own IV (Initial vector)
         *
         * @return the IV depending on the main digest method
         */
        public byte[] GetHashMagic()
        {
            // see https://www.ietf.org/rfc/rfc3110.txt
            // RSA/SHA1 SIG Resource Records
            byte[] result;
            switch (GetDigestAlgo().jceId)
            {
                case "sha1":
                    result = new byte[]
             { 0x30, 0x1f, 0x30, 0x07, 0x06, 0x05, 0x2b, 0x0e
            , 0x03, 0x02, 0x1a, 0x04, 0x14 };
                    break;
                case "sha224":
                    result = new byte[]
           { 0x30, 0x2b, 0x30, 0x0b, 0x06, 0x09, 0x60, (byte) 0x86
            , 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x04, 0x04, 0x1c };
                    break;
                case "sha256":
                    result = new byte[]
           { 0x30, 0x2f, 0x30, 0x0b, 0x06, 0x09, 0x60, (byte) 0x86
            , 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01, 0x04, 0x20 };
                    break;
                case "sha384":
                    result = new byte[]
           { 0x30, 0x3f, 0x30, 0x0b, 0x06, 0x09, 0x60, (byte) 0x86
            , 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x02, 0x04, 0x30 };
                    break;
                case "sha512":
                    result = new byte[]
           { 0x30, 0x4f, 0x30, 0x0b, 0x06, 0x09, 0x60, (byte) 0x86
            , 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x03, 0x04, 0x40 };
                    break;
                case "ripemd128":
                    result = new byte[]
        { 0x30, 0x1b, 0x30, 0x07, 0x06, 0x05, 0x2b, 0x24
            , 0x03, 0x02, 0x02, 0x04, 0x10 };
                    break;
                case "ripemd160":
                    result = new byte[]
        { 0x30, 0x1f, 0x30, 0x07, 0x06, 0x05, 0x2b, 0x24
            , 0x03, 0x02, 0x01, 0x04, 0x14 };
                    break;
                // case ripemd256: result = new byte[]
                //    { 0x30, 0x2b, 0x30, 0x07, 0x06, 0x05, 0x2b, 0x24
                //    , 0x03, 0x02, 0x03, 0x04, 0x20 };
                //    break;
                default:
                    throw new EncryptedDocumentException("Hash algorithm "
               + GetDigestAlgo() + " not supported for signing.");
            }

            return result;
        }

        /**
         * @return the uri for the signature method, i.e. currently only rsa is
         * supported, so it's the rsa variant of the main digest
         */
        public String GetSignatureMethodUri()
        {
            switch (GetDigestAlgo().jceId)
            {
                case "sha1": return XMLSignature.ALGO_ID_SIGNATURE_RSA_SHA1;
                case "sha224": return XMLSignature.ALGO_ID_SIGNATURE_RSA_SHA224;
                case "sha256": return XMLSignature.ALGO_ID_SIGNATURE_RSA_SHA256;
                case "sha384": return XMLSignature.ALGO_ID_SIGNATURE_RSA_SHA384;
                case "sha512": return XMLSignature.ALGO_ID_SIGNATURE_RSA_SHA512;
                case "ripemd160": return XMLSignature.ALGO_ID_SIGNATURE_RSA_RIPEMD160;
                default:
                    throw new EncryptedDocumentException("Hash algorithm "
               + GetDigestAlgo() + " not supported for signing.");
            }
        }

        /**
         * @return the uri for the main digest
         */
        public String GetDigestMethodUri()
        {
            return GetDigestMethodUri(GetDigestAlgo());
        }

        /**
         * @param digestAlgo the digest algo, currently only sha* and ripemd160 is supported 
         * @return the uri for the given digest
         */
        public static String GetDigestMethodUri(HashAlgorithm digestAlgo)
        {
            switch (digestAlgo.jceId)
            {
                case "sha1": return "http://www.w3.org/2000/09/xmldsig#sha1";
                case "sha224": return "http://www.w3.org/2001/04/xmldsig-more#sha224";
                case "sha256": return "http://www.w3.org/2001/04/xmlenc#sha256";
                case "sha384": return "http://www.w3.org/2001/04/xmldsig-more#sha384";
                case "sha512": return "http://www.w3.org/2001/04/xmlenc#sha512";
                case "ripemd160": return "http://www.w3.org/2001/04/xmlenc#ripemd160";
                default:
                    throw new EncryptedDocumentException("Hash algorithm "
               + digestAlgo + " not supported for signing.");
            }
        }

        /**
         * @param signatureFactory the xml signature factory, saved as thread-local
         */
        //public void SetSignatureFactory(XMLSignatureFactory signatureFactory)
        //{
        //    this.signatureFactory.Value = (signatureFactory);
        //}

        /**
         * @return the xml signature factory (thread-local)
         */
        //public XMLSignatureFactory GetSignatureFactory()
        //{
        //    XMLSignatureFactory sigFac = signatureFactory.Value;
        //    if (sigFac == null)
        //    {
        //        sigFac = XMLSignatureFactory.GetInstance("DOM", GetProvider());
        //        SetSignatureFactory(sigFac);
        //    }
        //    return sigFac;
        //}

        /**
         * @param keyInfoFactory the key factory, saved as thread-local
         */
        //public void SetKeyInfoFactory(KeyInfoFactory keyInfoFactory)
        //{
        //    this.keyInfoFactory.Value = (keyInfoFactory);
        //}

        /**
         * @return the key factory (thread-local)
         */
        //public KeyInfoFactory GetKeyInfoFactory()
        //{
        //    KeyInfoFactory keyFac = keyInfoFactory.Value;
        //    if (keyFac == null)
        //    {
        //        keyFac = KeyInfoFactory.GetInstance("DOM", GetProvider());
        //        SetKeyInfoFactory(keyFac);
        //    }
        //    return keyFac;
        //}

        /**
         * This method tests the existence of xml signature provider in the following order:
         * <ul>
         * <li>the class pointed to by the system property "jsr105Provider"</li>
         * <li>the Santuario xmlsec provider</li>
         * <li>the JDK xmlsec provider</li>
         * </ul>
         * 
         * For signing the classes are linked against the Santuario xmlsec, so this might
         * only work for validation (not tested).
         *  
         * @return the xml dsig provider
         */
        //public Provider GetProvider()
        //{
        //    Provider prov = provider.Get();
        //    if (prov == null)
        //    {
        //        String[] dsigProviderNames = {
        //        GetEnvironmentVariable("jsr105Provider"),
        //        "org.apache.jcp.xml.dsig.internal.dom.XMLDSigRI", // Santuario xmlsec
        //        "org.jcp.xml.dsig.internal.dom.XMLDSigRI"         // JDK xmlsec
        //    };
        //        foreach (String pn in dsigProviderNames)
        //        {
        //            if (pn == null) continue;
        //            try
        //            {
        //                prov = (Provider)Class.ForName(pn).newInstance();
        //                break;
        //            }
        //            catch (Exception e)
        //            {
        //                LOG.Log(POILogger.DEBUG, "XMLDsig-Provider '" + pn + "' can't be found - trying next.");
        //            }
        //        }
        //    }

        //    if (prov == null)
        //    {
        //        throw new Exception("JRE doesn't support default xml signature provider - Set jsr105Provider system property!");
        //    }

        //    return prov;
        //}

        /**
         * @return the cannonicalization method for XAdES-XL signing.
         * Defaults to <code>EXCLUSIVE</code>
         * @see <a href="http://docs.oracle.com/javase/7/docs/api/javax/xml/crypto/dsig/CanonicalizationMethod.html">javax.xml.Crypto.dsig.CanonicalizationMethod</a>
         */
        public String GetXadesCanonicalizationMethod()
        {
            return xadesCanonicalizationMethod;
        }

        /**
         * @param xadesCanonicalizationMethod the cannonicalization method for XAdES-XL signing
         * @see <a href="http://docs.oracle.com/javase/7/docs/api/javax/xml/crypto/dsig/CanonicalizationMethod.html">javax.xml.Crypto.dsig.CanonicalizationMethod</a>
         */
        public void SetXadesCanonicalizationMethod(String xadesCanonicalizationMethod)
        {
            this.xadesCanonicalizationMethod = xadesCanonicalizationMethod;
        }
    }
    public interface IData
    {

    }
    /**
 * Identifies a data object via a URI-Reference, as specified by
 * <a href="http://www.ietf.org/rfc/rfc2396.txt">RFC 2396</a>.
 *
 * <p>Note that some subclasses may not have a <code>type</code> attribute
 * and for objects of those types, the {@link #getType} method always returns
 * <code>null</code>.
 *
 * @author Sean Mullan
 * @author JSR 105 Expert Group
 * @since 1.6
 * @see URIDereferencer
 */
    public interface IURIReference
    {

        /**
         * Returns the URI of the referenced data object.
         *
         * @return the URI of the data object in RFC 2396 format (may be
         *    <code>null</code> if not specified)
         */
        String getURI();

        /**
         * Returns the type of data referenced by this URI.
         *
         * @return the type (a URI) of the data object (may be <code>null</code>
         *    if not specified)
         */
        String getType();
    }
    /**
 * A dereferencer of {@link URIReference}s.
 * <p>
 * The result of dereferencing a <code>URIReference</code> is either an
 * instance of {@link OctetStreamData} or {@link NodeSetData}. Unless the
 * <code>URIReference</code> is a <i>same-document reference</i> as defined
 * in section 4.2 of the W3C Recommendation for XML-Signature Syntax and
 * Processing, the result of dereferencing the <code>URIReference</code>
 * MUST be an <code>OctetStreamData</code>.
 *
 * @author Sean Mullan
 * @author Joyce Leung
 * @author JSR 105 Expert Group
 * @since 1.6
 * @see XMLCryptoContext#setURIDereferencer(URIDereferencer)
 * @see XMLCryptoContext#getURIDereferencer
 */
    public interface IURIDereferencer
    {

        /**
         * Dereferences the specified <code>URIReference</code> and returns the
         * dereferenced data.
         *
         * @param uriReference the <code>URIReference</code>
         * @param context an <code>XMLCryptoContext</code> that may
         *    contain additional useful information for dereferencing the URI. This
         *    implementation should dereference the specified
         *    <code>URIReference</code> against the context's <code>baseURI</code>
         *    parameter, if specified.
         * @return the dereferenced data
         * @throws NullPointerException if <code>uriReference</code> or
         *    <code>context</code> are <code>null</code>
         * @throws URIReferenceException if an exception occurs while
         *    dereferencing the specified <code>uriReference</code>
         */
        IData dereference(IURIReference uriReference, SignedXml context);
        //Data dereference(URIReference uriReference, XMLCryptoContext context)
    }
    public class XMLSignature
    {
        public static string ALGO_ID_MAC_HMAC_SHA1 = "http://www.w3.org/2000/09/xmldsig#hmac-sha1";
        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_DSA = "http://www.w3.org/2000/09/xmldsig#dsa-sha1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_DSA_SHA256 = "http://www.w3.org/2009/xmldsig11#dsa-sha256";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA1 = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_NOT_RECOMMENDED_RSA_MD5 = "http://www.w3.org/2001/04/xmldsig-more#rsa-md5";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_RIPEMD160 = "http://www.w3.org/2001/04/xmldsig-more#rsa-ripemd160";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA224 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha224";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA256 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA384 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha384";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA512 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA1_MGF1 = "http://www.w3.org/2007/05/xmldsig-more#sha1-rsa-MGF1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA224_MGF1 = "http://www.w3.org/2007/05/xmldsig-more#sha224-rsa-MGF1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA256_MGF1 = "http://www.w3.org/2007/05/xmldsig-more#sha256-rsa-MGF1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA384_MGF1 = "http://www.w3.org/2007/05/xmldsig-more#sha384-rsa-MGF1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_RSA_SHA512_MGF1 = "http://www.w3.org/2007/05/xmldsig-more#sha512-rsa-MGF1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_MAC_HMAC_NOT_RECOMMENDED_MD5 = "http://www.w3.org/2001/04/xmldsig-more#hmac-md5";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_MAC_HMAC_RIPEMD160 = "http://www.w3.org/2001/04/xmldsig-more#hmac-ripemd160";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_MAC_HMAC_SHA224 = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha224";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_MAC_HMAC_SHA256 = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha256";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_MAC_HMAC_SHA384 = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha384";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_MAC_HMAC_SHA512 = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha512";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_ECDSA_SHA1 = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha1";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_ECDSA_SHA224 = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha224";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_ECDSA_SHA256 = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_ECDSA_SHA384 = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha384";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_ECDSA_SHA512 = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha512";

        // Field descriptor #144 Ljava/lang/String;
        public static string ALGO_ID_SIGNATURE_ECDSA_RIPEMD160 = "http://www.w3.org/2007/05/xmldsig-more#ecdsa-ripemd160";
    }
    public interface IEvent
    {

    }
    /**
 *  The <code>EventListener</code> interface is the primary method for
 * handling events. Users implement the <code>EventListener</code> interface
 * and register their listener on an <code>EventTarget</code> using the
 * <code>AddEventListener</code> method. The users should also remove their
 * <code>EventListener</code> from its <code>EventTarget</code> after they
 * have completed using the listener.
 * <p> When a <code>Node</code> is copied using the <code>cloneNode</code>
 * method the <code>EventListener</code>s attached to the source
 * <code>Node</code> are not attached to the copied <code>Node</code>. If
 * the user wishes the same <code>EventListener</code>s to be added to the
 * newly created copy the user must add them manually.
 * <p>See also the <a href='http://www.w3.org/TR/2000/REC-DOM-Level-2-Events-20001113'>Document Object Model (DOM) Level 2 Events Specification</a>.
 * @since DOM Level 2
 */
    public interface IEventListener
    {
        /**
         *  This method is called whenever an event occurs of the type for which
         * the <code> EventListener</code> interface was registered.
         * @param evt  The <code>Event</code> contains contextual information
         *   about the event. It also contains the <code>stopPropagation</code>
         *   and <code>preventDefault</code> methods which are used in
         *   determining the event's flow and default action.
         */
        void handleEvent(IEvent evt);

    }
}
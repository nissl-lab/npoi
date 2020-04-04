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
    using NPOI.OpenXml4Net.OPC;
    using System;
    using System.Collections.Generic;
    using System.Security;
    using System.Security.Cryptography.Xml;
    using System.Xml;
    using System.Xml.Schema;

    /**
     * JSR105 Signature Facet base class.
     */
    public abstract class SignatureFacet : ISignatureConfigurable {

        //private static POILogger LOG = POILogFactory.GetLogger(typeof(SignatureFacet));

        public static String XML_NS = "http://www.w3.org/2000/xmlns/";
        public static String XML_DIGSIG_NS = SignedXml.XmlDsigNamespaceUrl;
        public static String OO_DIGSIG_NS = PackageNamespaces.DIGITAL_SIGNATURE;
        public static String MS_DIGSIG_NS = "http://schemas.microsoft.com/office/2006/digsig";
        public static String XADES_132_NS = "http://uri.etsi.org/01903/v1.3.2#";
        public static String XADES_141_NS = "http://uri.etsi.org/01903/v1.4.1#";

        protected SignatureConfig signatureConfig;

        public void SetSignatureConfig(SignatureConfig signatureConfig) {
            this.signatureConfig = signatureConfig;
        }

        /**
         * This method is being invoked by the XML signature service engine during
         * pre-sign phase. Via this method a signature facet implementation can add
         * signature facets to an XML signature.
         * 
         * @param document the signature document to be used for imports
         * @param references list of reference defInitions
         * @param objects objects to be signed/included in the signature document
         * @throws XMLSignatureException
         */
        public virtual void preSign(
              XmlDocument document
            , List<Reference> references
            , List<XmlNode> objects
        ) {
            // empty
        }

        /**
         * This method is being invoked by the XML signature service engine during
         * the post-sign phase. Via this method a signature facet can extend the XML
         * signatures with for example key information.
         *
         * @param document the signature document to be modified
         * @throws MarshalException
         */
        public virtual void postSign(XmlDocument document) {
            // empty
        }

        ////protected XMLSignatureFactory GetSignatureFactory() {
        ////    return signatureConfig.SignatureFactory;
        ////}

        protected Transform newTransform(String canonicalizationMethod) {
            ///return newTransform(canonicalizationMethod, null);
            throw new NotImplementedException();
        }

        ////protected Transform newTransform(String canonicalizationMethod, TransformParameterSpec paramSpec)
        ////{
        ////    try {
        ////        return GetSignatureFactory().newTransform(canonicalizationMethod, paramSpec);
        ////    } catch (SecurityException e) {
        ////        throw new XMLSignatureException("unknown canonicalization method: " + canonicalizationMethod, e);
        ////    }
        ////}

        protected Reference newReference(String uri, List<Transform> transforms, String type, String id, byte[] digestValue)
        {
            return newReference(uri, transforms, type, id, digestValue, signatureConfig);
        }

        public static Reference newReference(
              String uri
            , List<Transform> transforms
            , String type
            , String id
            , byte[] digestValue
        , SignatureConfig signatureConfig)
        {
            // the references appear in the package signature or the package object
            // so we can use the default digest algorithm
            //String digestMethodUri = signatureConfig.DigestMethodUri;
            //XMLSignatureFactory sigFac = signatureConfig.SignatureFactory;
            //DigestMethod digestMethod;
            //try {
            //    digestMethod = sigFac.NewDigestMethod(digestMethodUri, null);
            //} catch (SecurityException e) {
            //    throw new XMLSignatureException("unknown digest method uri: " + digestMethodUri, e);
            //}

            //Reference reference;
            //if (digestValue == null) {
            //    reference = sigFac.NewReference(uri, digestMethod, transforms, type, id);
            //} else {
            //    reference = sigFac.NewReference(uri, digestMethod, transforms, type, id, digestValue);
            //}

            //brokenJvmWorkaround(reference);

            //return reference;
            throw new NotImplementedException();
        }

        // helper method ... will be Removed soon
        public static void brokenJvmWorkaround(Reference reference) {
            throw new NotImplementedException();
            //DigestMethod digestMethod = reference.DigestMethod;
            //String digestMethodUri = digestMethod.Algorithm;

            //// workaround for https://bugzilla.redhat.com/Show_bug.cgi?id=1155012
            //// overwrite standard message digest, if a digest <> SHA1 is used
            //Provider bcProv = Security.GetProvider("BC");
            //if (bcProv != null && !DigestMethod.SHA1.Equals(digestMethodUri)) {
            //    try {
            //        Method m = DOMDigestMethod.class.GetDeclaredMethod("getMessageDigestAlgorithm");
            //        m.Accessible=(/*setter*/true);
            //        String mdAlgo = (String)m.Invoke(digestMethod);
            //        MessageDigest md = MessageDigest.GetInstance(mdAlgo, bcProv);
            //        Field f = DOMReference.class.GetDeclaredField("md");
            //        f.Accessible=(/*setter*/true);
            //        f.Set(reference, md);
            //    } catch (Exception e) {
            //        LOG.Log(POILogger.WARN, "Can't overwrite message digest (workaround for https://bugzilla.redhat.com/Show_bug.cgi?id=1155012)", e);
            //    }
            //}
        }
    }
}
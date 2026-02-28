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
    using System.Xml;

    /**
* Signature Facet implementation that Adds ds:KeyInfo to the XML signature.
* 
* @author Frank Cornelis
* 
*/
    public class KeyInfoSignatureFacet : SignatureFacet
    {


        public override void postSign(XmlDocument document)
        {
            throw new NotImplementedException();

            //XmlNodeList nl = document.GetElementsByTagNameNS(XML_DIGSIG_NS, "Object");

            ///*
            // * Make sure we insert right After the ds:SignatureValue element, just
            // * before the first ds:Object element.
            // */
            //XmlNode nextSibling = (nl.Count == 0) ? null : nl.Item(0);

            ///*
            // * Construct the ds:KeyInfo element using JSR 105.
            // */
            //KeyInfoFactory keyInfoFactory = signatureConfig.KeyInfoFactory;
            //List<Object> x509DataObjects = new List<Object>();
            //X509Certificate signingCertificate = signatureConfig.SigningCertificateChain.Get(0);

            //List<Object> keyInfoContent = new List<Object>();

            //if (signatureConfig.IsIncludeKeyValue())
            //{
            //    KeyValue keyValue;
            //    try
            //    {
            //        keyValue = keyInfoFactory.NewKeyValue(signingCertificate.PublicKey);
            //    }
            //    catch (KeyException e)
            //    {
            //        throw new Exception("key exception: " + e.Message, e);
            //    }
            //    keyInfoContent.Add(keyValue);
            //}

            //if (signatureConfig.IsIncludeIssuerSerial())
            //{
            //    x509DataObjects.Add(keyInfoFactory.NewX509IssuerSerial(
            //        signingCertificate.IssuerX500Principal.ToString(),
            //        signingCertificate.SerialNumber));
            //}

            //if (signatureConfig.IsIncludeEntireCertificateChain())
            //{
            //    x509DataObjects.AddAll(signatureConfig.SigningCertificateChain);
            //}
            //else
            //{
            //    x509DataObjects.Add(signingCertificate);
            //}

            //if (!x509DataObjects.IsEmpty())
            //{
            //    X509Data x509Data = keyInfoFactory.NewX509Data(x509DataObjects);
            //    keyInfoContent.Add(x509Data);
            //}
            //KeyInfo keyInfo = keyInfoFactory.NewKeyInfo(keyInfoContent);
            //DOMKeyInfo domKeyInfo = (DOMKeyInfo)keyInfo;

            //Key key = new Key() {
            //    private static long serialVersionUID = 1L;

            //    public String GetAlgorithm() {
            //        return null;
            //    }

            //    public byte[] GetEncoded() {
            //        return null;
            //    }

            //    public String GetFormat() {
            //        return null;
            //    }
            //};

            //Element n = document.DocumentElement;
            //DOMSignContext domSignContext = new DOMSignContext(key, n, nextSibling);
            //foreach (Entry<String, String> me in signatureConfig.NamespacePrefixes.EntrySet())
            //{
            //    domSignContext.PutNamespacePrefix(me.Key, me.Value);
            //}

            //DOMStructure domStructure = new DOMStructure(n);
            //domKeyInfo.Marshal(domStructure, domSignContext);

            //// Move keyinfo into the right place
            //if (nextSibling != null)
            //{
            //    NodeList kiNl = document.GetElementsByTagNameNS(XML_DIGSIG_NS, "KeyInfo");
            //    if (kiNl.Length != 1)
            //    {
            //        throw new Exception("KeyInfo wasn't Set");
            //    }
            //    nextSibling.ParentNode.InsertBefore(kiNl.Item(0), nextSibling);
            //}
        }
    }
}
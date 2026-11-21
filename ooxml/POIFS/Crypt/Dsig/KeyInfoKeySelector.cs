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
    using Org.BouncyCastle.X509;
    using System;
    using System.Collections.Generic;

    public interface IKeySelectorResult
    {

        /**
         * Returns the selected key.
         *
         * @return the selected key, or <code>null</code> if none can be found
         */
        IKey GetKey();
    }
    /**
     * JSR105 key selector implementation using the ds:KeyInfo data of the signature
     * itself.
     */
    public class KeyInfoKeySelector : KeySelector, IKeySelectorResult
    {


        private List<X509Certificate> certChain = new List<X509Certificate>();

        public IKeySelectorResult select(KeyInfo keyInfo, Purpose purpose, AlgorithmMethod method, XMLCryptoContext context)
        {
            if (null == keyInfo)
            {
                throw new Exception("no ds:KeyInfo present");
            }
            List<XMLStructure> keyInfoContent = keyInfo.Content;
            certChain.Clear();
            foreach (XMLStructure keyInfoStructure in keyInfoContent)
            {
                if (!(keyInfoStructure is X509Data))
                {
                    continue;
                }
                X509Data x509Data = (X509Data)keyInfoStructure;
                List<Object> x509DataList = x509Data.Content;
                foreach (Object x509DataObject in x509DataList)
                {
                    if (!(x509DataObject is X509Certificate))
                    {
                        continue;
                    }
                    X509Certificate certificate = (X509Certificate)x509DataObject;
                    certChain.Add(certificate);
                }
            }
            if (certChain.Count == 0)
            {
                throw new Exception("No key found!");
            }
            return this;
        }

        public IKey GetKey()
        {
            // The first certificate is presumably the signer.
            //return certChain.Count == 0 ? null : certChain[0].GetPublicKey().;
            throw new NotImplementedException();
        }

        /**
         * Gives back the X509 certificate used during the last signature
         * verification operation.
         * 
         * @return the certificate which was used to sign the xml content
         */
        public X509Certificate GetSigner()
        {
            // The first certificate is presumably the signer.
            return certChain.Count == 0 ? null : certChain[(0)];
        }

        public List<X509Certificate> GetCertChain()
        {
            return certChain;
        }
    }

    internal sealed class X509Data
    {
        public List<object> Content { get; internal set; }
    }
}
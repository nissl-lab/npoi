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
    using System;
    using System.Collections.Generic;
    using System.Runtime.Serialization;

    /**
     * Container class for PKI revocation data.
     * 
     * @author Frank Cornelis
     * 
     */
    public class RevocationData
    {

        private List<byte[]> crls;

        private List<byte[]> ocsps;

        /**
         * Default constructor.
         */
        public RevocationData()
        {
            this.crls = new List<byte[]>();
            this.ocsps = new List<byte[]>();
        }

        /**
         * Adds a CRL to this revocation data Set.
         * 
         * @param encodedCrl
         */
        public void AddCRL(byte[] encodedCrl)
        {
            this.crls.Add(encodedCrl);
        }

        /**
         * Adds a CRL to this revocation data Set.
         * 
         * @param crl
         */
        public void AddCRL(X509CRL crl)
        {
            byte[] encodedCrl;
            try
            {
                encodedCrl = crl.getEncoded(); ;
            }
            catch (CRLException e)
            {
                throw new ArgumentException("CRL coding error: "
                        + e.Message, e);
            }
            AddCRL(encodedCrl);
        }

        /**
         * Adds an OCSP response to this revocation data Set.
         * 
         * @param encodedOcsp
         */
        public void AddOCSP(byte[] encodedOcsp)
        {
            this.ocsps.Add(encodedOcsp);
        }

        /**
         * Gives back a list of all CRLs.
         * 
         * @return a list of all CRLs
         */
        public List<byte[]> GetCRLs()
        {
            return this.crls;
        }

        /**
         * Gives back a list of all OCSP responses.
         * 
         * @return a list of all OCSP response
         */
        public List<byte[]> GetOCSPs()
        {
            return this.ocsps;
        }

        /**
         * Returns <code>true</code> if this revocation data Set holds OCSP
         * responses.
         * 
         * @return <code>true</code> if this revocation data Set holds OCSP
         * responses.
         */
        public bool HasOCSPs()
        {
            return this.ocsps.Count > 0;
        }

        /**
         * Returns <code>true</code> if this revocation data Set holds CRLs.
         * 
         * @return <code>true</code> if this revocation data Set holds CRLs.
         */
        public bool HasCRLs()
        {
            return this.crls.Count > 0;
        }

        /**
         * Returns <code>true</code> if this revocation data is not empty.
         * 
         * @return <code>true</code> if this revocation data is not empty.
         */
        public bool HasRevocationDataEntries()
        {
            return HasOCSPs() || HasCRLs();
        }
    }

    public class X509CRL
    {
        public byte[] getEncoded()
        {
            throw new NotImplementedException();
        }
    }

    [Serializable]
    internal sealed class CRLException : Exception
    {
        public CRLException()
        {
        }

        public CRLException(string message) : base(message)
        {
        }

        public CRLException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected CRLException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
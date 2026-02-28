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
    using System.Security.Cryptography.X509Certificates;




    /**
     * Interface for a service that retrieves revocation data about some given
     * certificate chain.
     * 
     * @author Frank Cornelis
     * 
     */
    public interface IRevocationDataService
    {

        /**
         * Gives back the revocation data corresponding with the given certificate
         * chain.
         * 
         * @param certificateChain the certificate chain
         * @return the revocation data corresponding with the given certificate chain.
         */
        RevocationData GetRevocationData(List<X509Certificate> certificateChain);
    }

}
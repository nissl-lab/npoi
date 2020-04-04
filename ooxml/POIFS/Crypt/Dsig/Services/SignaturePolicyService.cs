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

    /**
     * Interface for the signature policy service.
     * 
     * @author Frank Cornelis
     * 
     */
    public interface ISignaturePolicyService
    {

        /**
         * Gives back the signature policy identifier URI.
         * 
         * @return the signature policy identifier URI.
         */
        String GetSignaturePolicyIdentifier();

        /**
         * Gives back the short description of the signature policy or
         * <code>null</code> if a description is not available.
         * 
         * @return the description, or <code>null</code>.
         */
        String GetSignaturePolicyDescription();

        /**
         * Gives back the download URL where the signature policy document can be
         * found. Can be <code>null</code> in case such a download location does not
         * exist.
         * 
         * @return the download URL, or <code>null</code>.
         */
        String GetSignaturePolicyDownloadUrl();

        /**
         * Gives back the signature policy document.
         * 
         * @return the bytes of the signature policy document.
         */
        byte[] GetSignaturePolicyDocument();
    }
}
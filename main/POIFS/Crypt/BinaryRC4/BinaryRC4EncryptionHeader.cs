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

namespace NPOI.POIFS.Crypt.BinaryRC4
{
    using System;

    using NPOI.POIFS.Crypt;

    using NPOI.POIFS.Crypt.Standard;
    using NPOI.Util;

    public class BinaryRC4EncryptionHeader : EncryptionHeader, EncryptionRecord
    {

        protected internal BinaryRC4EncryptionHeader()
        {
            CipherAlgorithm = (CipherAlgorithm.rc4);
            KeySize = (40);
            BlockSize = (-1);
            CipherProvider = (CipherProvider.rc4);
            HashAlgorithm = (HashAlgorithm.md5);
            SizeExtra = (0);
            Flags = (0);
            CspName = ("");
            ChainingMode = (null);
        }

        public void Write(LittleEndianByteArrayOutputStream littleendianbytearrayoutputstream)
        {
        }
    }

}
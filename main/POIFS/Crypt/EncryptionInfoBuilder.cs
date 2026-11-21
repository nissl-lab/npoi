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

using NPOI.Util;

namespace NPOI.POIFS.Crypt
{
    using System;


    public interface IEncryptionInfoBuilder
    {
        /**
     * Initialize the builder from a stream
     */
        void Initialize(EncryptionInfo ei, ILittleEndianInput dis);

        /**
     * Initialize the builder from scratch
     */

        void Initialize(EncryptionInfo ei, CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm, int keyBits,
            int blockSize, ChainingMode chainingMode);

        /**
     * @return the header data
     */
        EncryptionHeader GetHeader();

        /**
     * @return the verifier data
     */
        EncryptionVerifier GetVerifier();

        /**
     * @return the decryptor
     */
        Decryptor GetDecryptor();

        /**
     * @return the encryptor
     */
        Encryptor GetEncryptor();

        /**
		* @return the EncryptionInfo
		*/
        EncryptionInfo GetEncryptionInfo();
    }

}
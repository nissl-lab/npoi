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

namespace NPOI.POIFS.Crypt
{
    using System;

    /**
     * Office supports various encryption modes.
     * The encryption is either based on the whole Container ({@link #agile}, {@link #standard} or {@link #binaryRC4})
     * or record based ({@link #cryptoAPI}). The record based encryption can't be accessed directly, but will be
     * invoked by using the {@link Biff8EncryptionKey#setCurrentUserPassword(String)} before saving the document.
     */
    public class EncryptionMode
    {
        /* @see <a href="http://msdn.microsoft.com/en-us/library/dd907466(v=office.12).aspx">2.3.6 Office Binary Document RC4 Encryption</a> */
        public static readonly EncryptionMode BinaryRC4 = new EncryptionMode("NPOI.POIFS.Crypt.BinaryRC4.BinaryRC4EncryptionInfoBuilder", 1, 1, 0x0);
        /* @see <a href="http://msdn.microsoft.com/en-us/library/dd905225(v=office.12).aspx">2.3.5 Office Binary Document RC4 CryptoAPI Encryption</a> */
        public static readonly EncryptionMode CryptoAPI = new EncryptionMode("NPOI.POIFS.Crypt.CryptoAPI.CryptoAPIEncryptionInfoBuilder", 4, 2, 0x04);
        /* @see <a href="http://msdn.microsoft.com/en-us/library/dd906097(v=office.12).aspx">2.3.4.5 \EncryptionInfo Stream (Standard Encryption)</a> */
        public static readonly EncryptionMode Standard = new EncryptionMode("NPOI.POIFS.Crypt.Standard.StandardEncryptionInfoBuilder", 4, 2, 0x24);
        /* @see <a href="http://msdn.microsoft.com/en-us/library/dd925810(v=office.12).aspx">2.3.4.10 \EncryptionInfo Stream (Agile Encryption)</a> */
        public static readonly EncryptionMode Agile = new EncryptionMode("NPOI.POIFS.Crypt.Agile.AgileEncryptionInfoBuilder", 4, 4, 0x40);


        public string Builder { get; private set; }
        public int VersionMajor { get;private set; }
        public int VersionMinor { get;private set; }
        public int EncryptionFlags { get;private set; }

        public EncryptionMode(string builder, int versionMajor, int versionMinor, int encryptionFlags)
        {
            this.Builder = builder;
            this.VersionMajor = versionMajor;
            this.VersionMinor = versionMinor;
            this.EncryptionFlags = encryptionFlags;
        }
    }

}
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

using System.IO;
using NPOI.POIFS.FileSystem;
using NPOI.Util;

namespace NPOI.POIFS.Crypt
{
    using System;
    using System.Reflection;

    public class EncryptionInfo
    {

        /**
         * Document entry name for encryption info xml descriptor
         */
        public const String ENCRYPTION_INFO_ENTRY = "EncryptionInfo";

        /**
     * A flag that specifies whether CryptoAPI RC4 or ECMA-376 encryption
     * ECMA-376 is used. It MUST be 1 unless flagExternal is 1. If flagExternal is 1, it MUST be 0.
     */
        public static BitField flagCryptoAPI = BitFieldFactory.GetInstance(0x04);

        /**
     * A value that MUST be 0 if document properties are encrypted.
     * The encryption of document properties is specified in section 2.3.5.4.
     */
        public static BitField flagDocProps = BitFieldFactory.GetInstance(0x08);

        /**
     * A value that MUST be 1 if extensible encryption is used. If this value is 1,
     * the value of every other field in this structure MUST be 0.
     */
        public static BitField flagExternal = BitFieldFactory.GetInstance(0x10);

        /**
     * A value that MUST be 1 if the protected content is an ECMA-376 document
     * ECMA-376. If the fAES bit is 1, the fCryptoAPI bit MUST also be 1.
     */
        public static BitField flagAES = BitFieldFactory.GetInstance(0x20);


        /**
     * Opens for decryption
     */

        public EncryptionInfo(POIFSFileSystem fs)
            : this(fs.Root)
        {
        }

        /**
     * Opens for decryption
     */
        public EncryptionInfo(OPOIFSFileSystem fs)
            : this(fs.Root)
        {
        }
        /**
     * Opens for decryption
     */

        public EncryptionInfo(NPOIFSFileSystem fs) : this(fs.Root)
        {

        }

        /**
     * Opens for decryption
     */

        public EncryptionInfo(DirectoryNode dir)
            : this(dir.CreateDocumentInputStream("EncryptionInfo"), null)
        {

        }

        public EncryptionInfo(ILittleEndianInput dis, EncryptionMode preferredEncryptionMode)
        {
            if (preferredEncryptionMode == EncryptionMode.XOR)
            {
                _versionMajor = EncryptionMode.XOR.VersionMajor;
                _versionMinor = EncryptionMode.XOR.VersionMinor;
            }
            else
            {
                _versionMajor = dis.ReadShort();
                _versionMinor = dis.ReadShort();
            }

            if (VersionMajor == EncryptionMode.XOR.VersionMajor
               && VersionMinor == EncryptionMode.XOR.VersionMinor)
            {
                _encryptionMode = EncryptionMode.XOR;
                _encryptionFlags = -1;
            }

            else if (VersionMajor == EncryptionMode.BinaryRC4.VersionMajor
                 && VersionMinor == EncryptionMode.BinaryRC4.VersionMinor)
            {
                _encryptionMode = EncryptionMode.BinaryRC4;
                _encryptionFlags = -1;
            }
            else if (2 <= VersionMajor && VersionMajor <= 4 && VersionMinor == 2)
            {
                _encryptionFlags = dis.ReadInt();
                _encryptionMode = (preferredEncryptionMode == EncryptionMode.CryptoAPI || !flagAES.IsSet(_encryptionFlags)) ? EncryptionMode.CryptoAPI : EncryptionMode.Standard;
            }
            else if (VersionMajor == EncryptionMode.Agile.VersionMajor && VersionMinor == EncryptionMode.Agile.VersionMinor)
            {
                _encryptionMode = EncryptionMode.Agile;
                _encryptionFlags = dis.ReadInt();
            }
            else
            {
                _encryptionFlags = dis.ReadInt();
                throw new EncryptedDocumentException(
                    "Unknown encryption: version major: " + VersionMajor +
                    " / version minor: " + VersionMinor +
                    " / fCrypto: " + flagCryptoAPI.IsSet(EncryptionFlags) +
                    " / fExternal: " + flagExternal.IsSet(EncryptionFlags) +
                    " / fDocProps: " + flagDocProps.IsSet(EncryptionFlags) +
                    " / fAES: " + flagAES.IsSet(EncryptionFlags));
            }

            IEncryptionInfoBuilder eib;
            try
            {
                eib = GetBuilder(_encryptionMode);
            }
            catch (Exception e)
            {
                throw new IOException(e.Message, e);
            }

            eib.Initialize(this, dis);
            _header = eib.GetHeader();
            _verifier = eib.GetVerifier();
            _decryptor = eib.GetDecryptor();
            _encryptor = eib.GetEncryptor();
        }

        /**
     * @deprecated Use {@link #EncryptionInfo(EncryptionMode)} (fs parameter no longer required)
     */

        public EncryptionInfo(POIFSFileSystem fs, EncryptionMode encryptionMode)
            : this(encryptionMode)
        {
            ;
        }

        /**
     * @deprecated Use {@link #EncryptionInfo(EncryptionMode)} (fs parameter no longer required)
     */

        public EncryptionInfo(NPOIFSFileSystem fs, EncryptionMode encryptionMode)
            : this(encryptionMode)
        {
            ;
        }

        /**
     * @deprecated Use {@link #EncryptionInfo(EncryptionMode)} (dir parameter no longer required)
     */

        public EncryptionInfo(DirectoryNode dir, EncryptionMode encryptionMode)
            : this(encryptionMode)
        {
            ;
        }

        /**
     * @deprecated use {@link #EncryptionInfo(EncryptionMode, CipherAlgorithm, HashAlgorithm, int, int, ChainingMode)}
     */

        public EncryptionInfo(
            POIFSFileSystem fs
            , EncryptionMode encryptionMode
            , CipherAlgorithm cipherAlgorithm
            , HashAlgorithm hashAlgorithm
            , int keyBits
            , int blockSize
            , ChainingMode chainingMode
            )
            : this(encryptionMode, cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode)
        {
            ;
        }

        /**
     * @deprecated use {@link #EncryptionInfo(EncryptionMode, CipherAlgorithm, HashAlgorithm, int, int, ChainingMode)}
     */

        public EncryptionInfo(
            NPOIFSFileSystem fs
            , EncryptionMode encryptionMode
            , CipherAlgorithm cipherAlgorithm
            , HashAlgorithm hashAlgorithm
            , int keyBits
            , int blockSize
            , ChainingMode chainingMode
            )
            : this(encryptionMode, cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode)
        {
            ;
        }

        /**
     * @deprecated use {@link #EncryptionInfo(EncryptionMode, CipherAlgorithm, HashAlgorithm, int, int, ChainingMode)}
     */

        public EncryptionInfo(
            DirectoryNode dir
            , EncryptionMode encryptionMode
            , CipherAlgorithm cipherAlgorithm
            , HashAlgorithm hashAlgorithm
            , int keyBits
            , int blockSize
            , ChainingMode chainingMode
            )
            : this(encryptionMode, cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode)
        {
            ;
        }

        /**
     * Prepares for encryption, using the given Encryption Mode, and
     *  all other parameters as default.
     * @see #EncryptionInfo(EncryptionMode, CipherAlgorithm, HashAlgorithm, int, int, ChainingMode)
     */

        public EncryptionInfo(EncryptionMode encryptionMode)
            : this(encryptionMode, null, null, -1, -1, null)
        {

        }

        /**
     * Constructs an EncryptionInfo from scratch
     *
     * @param encryptionMode see {@link EncryptionMode} for values, {@link EncryptionMode#cryptoAPI} is for
     *   internal use only, as it's record based
     * @param cipherAlgorithm
     * @param hashAlgorithm
     * @param keyBits
     * @param blockSize
     * @param chainingMode
     * 
     * @throws EncryptedDocumentException if the given parameters mismatch, e.g. only certain combinations
     *   of keyBits, blockSize are allowed for a given {@link CipherAlgorithm}
     */

        public EncryptionInfo(
            EncryptionMode encryptionMode
            , CipherAlgorithm cipherAlgorithm
            , HashAlgorithm hashAlgorithm
            , int keyBits
            , int blockSize
            , ChainingMode chainingMode
            )
        {
            _encryptionMode = encryptionMode;
            _versionMajor = encryptionMode.VersionMajor;
            _versionMinor = encryptionMode.VersionMinor;
            _encryptionFlags = encryptionMode.EncryptionFlags;

            IEncryptionInfoBuilder eib;
            try
            {
                eib = GetBuilder(encryptionMode);
            }
            catch(Exception e)
            {
                throw new EncryptedDocumentException(e);
            }

            eib.Initialize(this, cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode);

            _header = eib.GetHeader();
            _verifier = eib.GetVerifier();
            _decryptor = eib.GetDecryptor();
            _encryptor = eib.GetEncryptor();
        }

        protected static IEncryptionInfoBuilder GetBuilder(EncryptionMode encryptionMode)
        {
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
            Type t = null;
            foreach(Assembly assembly in assemblies)
            {
                t = assembly.GetType(encryptionMode.Builder);
                if(t != null)
                    break;
            }
            if(t == null)
            {
                throw new EncryptedDocumentException("Not found type " + encryptionMode.Builder);
            }
            IEncryptionInfoBuilder eib = null;
            eib = (IEncryptionInfoBuilder) t.Assembly.CreateInstance(encryptionMode.Builder);
            return eib;
        }

        private readonly int _versionMajor;
        public int VersionMajor
        {
            get { return _versionMajor; }
        }

        private readonly int _versionMinor;
        public int VersionMinor
        {
            get { return _versionMinor; }
        }

        private readonly int _encryptionFlags;
        public int EncryptionFlags
        {
            get { return _encryptionFlags; }
        }
        private readonly EncryptionHeader _header;
        public EncryptionHeader Header
        {
            get { return _header; }
        }
        private readonly EncryptionVerifier _verifier;
        public EncryptionVerifier Verifier
        {
            get { return _verifier; }
        }
        private readonly Decryptor _decryptor;
        public Decryptor Decryptor
        {
            get { return _decryptor; }
        }
        private readonly Encryptor _encryptor;
        public Encryptor Encryptor
        {
            get { return _encryptor; }
        }

        private readonly EncryptionMode _encryptionMode;
        public EncryptionMode EncryptionMode
        {
            get { return _encryptionMode; }
        }

        /**
 * @return true, if Document Summary / Summary are encrypted and stored in the {@code EncryptedStream} stream,
 * otherwise the Summaries aren't encrypted and located in their usual streams
 */
        public bool IsDocPropsEncrypted()
        {
            return !flagDocProps.IsSet(_encryptionFlags);
        }

        public override string ToString()
        {
            return "EncryptionInfo:" +
                   "\n\tversionMajor: " + VersionMajor +
                   "\n\tversionMinor: " + VersionMinor +
                   "\n\tencryptionFlags: " + EncryptionFlags +
                   "\n\theader: " + Header +
                   "\n\tverifier: " + Verifier +
                   "\n\tdecryptor: " + Decryptor +
                   "\n\tencryptor: " + Encryptor;
        }
    }
}
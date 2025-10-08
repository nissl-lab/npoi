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

namespace NPOI.POIFS.Crypt.Agile
{
    using System;
    using System.IO;
    using System.Xml.Linq;
    using FileSystem;

    /// <summary>
    /// Agile Encryption implementation that uses XlsxEncryptor helper methods.
    /// This class provides POI-compatible API for encrypting XLSX files.
    /// </summary>
    public class AgileEncryptorForXlsx : Encryptor
    {
        #nullable enable
        private byte[]? _integritySalt;
        private byte[]? _encryptionKey;
        private byte[]? _keySalt;
        private XDocument? _encryptionInfoXml;
        private string _password = "";

        public AgileEncryptorForXlsx()
        {
        }

        protected AgileEncryptorForXlsx(AgileEncryptorForXlsx other)
        {
            _integritySalt = other._integritySalt?.Clone() as byte[];
            _encryptionKey = other._encryptionKey?.Clone() as byte[];
            _keySalt = other._keySalt?.Clone() as byte[];
            _encryptionInfoXml = other._encryptionInfoXml != null
                ? new XDocument(other._encryptionInfoXml)
                : null;
        }

        /// <summary>
        /// Confirms the password and generates encryption parameters.
        /// This is equivalent to Apache POI's confirmPassword method.
        /// </summary>
        public override void ConfirmPassword(string password)
        {
            if (string.IsNullOrEmpty(password))
            {
                throw new ArgumentException("Password cannot be null or empty", nameof(password));
            }

            if (password.Length > 255)
            {
                throw new ArgumentException("Password is too long (max 255 characters)", nameof(password));
            }

            this._password = password;

            var (xmlDoc, encryptionKey, keySalt, integritySalt) =
                XlsxEncryptor.GenerateEncryptionInfo(password);

            _encryptionInfoXml = xmlDoc;
            _encryptionKey = encryptionKey;
            _keySalt = keySalt;
            _integritySalt = integritySalt;
        }

        /// <summary>
        /// For compatibility with existing code that passes all parameters.
        /// </summary>
        public override void ConfirmPassword(
            string password,
            byte[] keySpec,
            byte[] keySalt,
            byte[] verifier,
            byte[] verifierSalt,
            byte[] integritySalt)
        {
            ConfirmPassword(password);
        }

        /// <summary>
        /// Gets the output stream for writing encrypted data.
        /// This is equivalent to Apache POI's getDataStream(DirectoryNode) method.
        /// </summary>
        public override OutputStream GetDataStream(DirectoryNode dir)
        {
            if (_encryptionKey == null || _keySalt == null || _integritySalt == null || _encryptionInfoXml == null)
            {
                throw new InvalidOperationException(
                    "Password must be confirmed before getting data stream. Call ConfirmPassword first.");
            }

            return new AgileCipherOutputStream(dir, this);
        }

        /// <summary>
        /// Internal stream implementation that handles the encryption process.
        /// </summary>
        public class AgileCipherOutputStream(DirectoryNode dir, AgileEncryptorForXlsx encryptor) : OutputStream
        {
            private readonly MemoryStream _buffer = new();
            private readonly DirectoryNode _dir = dir ?? throw new ArgumentNullException(nameof(dir));
            private readonly AgileEncryptorForXlsx _encryptor = encryptor ?? throw new ArgumentNullException(nameof(encryptor));
            private bool _isClosed;
            private bool _isDisposd;

            public string GetPassword() => _encryptor._password;

            public override bool CanRead => false;
            public override bool CanSeek => true;
            public override bool CanWrite => true;
            public override long Length => _buffer.Length;

            public override long Position
            {
                get => _buffer.Position;
                set => _buffer.Position = value;
            }

            public override void Write(int b)
            {
                Console.WriteLine("Warning: Write(int) is inefficient, consider using Write(byte[], int, int) instead.");
            }



            public override void Write(byte[] b, int off, int len)
            {
                _buffer.Write(b, off, len);
            }

            public override void Close()
            {
                byte[] zipBytes = _buffer.ToArray();

                if(_isClosed)
                {
                    return;
                }
                _isClosed = true;
                byte[] encryptedBytes = XlsxEncryptor.FromBytesToBytes(zipBytes, _encryptor._password);
                _dir.NFileSystem.Data.Write(new ByteBuffer(encryptedBytes, 0, encryptedBytes.Length), 0);

                _dir.FileSystem.MarkAsDirectWrite();

                base.Close();
            }

            public override void Flush()
            {
                _buffer.Flush();
            }

            protected override void Dispose(bool disposing)
            {
                if (!_isDisposd && disposing)
                    base.Dispose(disposing);
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                throw new NotSupportedException("Cannot read from encryption stream");
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                return _buffer.Seek(offset, origin);
            }

            public override void SetLength(long value)
            {
                _buffer.SetLength(value);
            }
        }
    }
}
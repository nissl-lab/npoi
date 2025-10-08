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

            // Use XlsxEncryptor helper to generate all encryption parameters
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
            // For now, just use the simple version
            // This could be extended to use the provided parameters
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
        /// Creates a copy of this encryptor.
        /// </summary>
        public Encryptor Copy()
        {
            return new AgileEncryptorForXlsx(this);
        }

        /// <summary>
        /// Internal stream implementation that handles the encryption process.
        /// </summary>
        private sealed class AgileCipherOutputStream(DirectoryNode dir, AgileEncryptorForXlsx encryptor) : OutputStream
        {
            private readonly MemoryStream _buffer = new();
            private readonly DirectoryNode _dir = dir ?? throw new ArgumentNullException(nameof(dir));
            private readonly AgileEncryptorForXlsx _encryptor = encryptor ?? throw new ArgumentNullException(nameof(encryptor));
            private bool _disposed;

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
            }

#pragma warning disable CA1513
            public override void Write(byte[] b, int off, int len)
            {
                if (_disposed)
                    throw new ObjectDisposedException(nameof(AgileCipherOutputStream));

                _buffer.Write(b, off, len);
            }
#pragma warning restore CA1513

            public override void Flush()
            {
                _buffer.Flush();
            }

            protected override void Dispose(bool disposing)
            {
                if (!_disposed && disposing)
                {
                    try
                    {
                        // Get the unencrypted data
                        byte[] packageData = _buffer.ToArray();

                        if (packageData.Length == 0)
                        {
                            throw new InvalidOperationException("No data written to encryption stream");
                        }

                        // Encrypt the package using XlsxEncryptor helper
                        byte[] encryptedPackage = XlsxEncryptor.EncryptPackage(
                            packageData,
                            _encryptor._encryptionKey!,
                            _encryptor._keySalt!
                        );

                        // Update integrity HMAC
                        XlsxEncryptor.UpdateIntegrityHmac(
                            encryptedPackage,
                            packageData.Length,
                            _encryptor._encryptionKey!,
                            _encryptor._keySalt!,
                            _encryptor._integritySalt!,
                            _encryptor._encryptionInfoXml!
                        );

                        // Create EncryptedPackage stream in the directory
                        using (var encPackageMs = new MemoryStream(encryptedPackage))
                        {
                            _dir.CreateDocument("EncryptedPackage", encPackageMs);
                        }

                        // Create EncryptionInfo stream
                        CreateEncryptionInfoStream(_dir, _encryptor._encryptionInfoXml!);

                        // Create DataSpaces structure
                        CreateDataSpacesStructure(_dir);
                    }
                    finally
                    {
                        _buffer.Dispose();
                        _disposed = true;
                    }
                }

                base.Dispose(disposing);
            }

            private static void CreateEncryptionInfoStream(DirectoryNode dir, XDocument encryptionInfo)
            {
                using var ms = new MemoryStream();
                using (var bw = new BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    // Write version info
                    bw.Write((ushort)4);  // version major
                    bw.Write((ushort)4);  // version minor
                    bw.Write((uint)0x40); // flags

                    // Write XML
                    if (encryptionInfo.Root != null)
                    {
                        string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                     encryptionInfo.Root.ToString(SaveOptions.DisableFormatting);
                        xml = xml.Replace(" />", "/>");
                        byte[] xmlBytes = System.Text.Encoding.UTF8.GetBytes(xml);
                        bw.Write(xmlBytes);
                    }
                }

                ms.Position = 0;
                dir.CreateDocument("EncryptionInfo", ms);
            }

            private static void CreateDataSpacesStructure(DirectoryNode rootDir)
            {
                // Create \u0006DataSpaces directory
                var dataSpacesDir = rootDir.CreateDirectory("\u0006DataSpaces");

                // Create Version stream
                CreateVersionStream(dataSpacesDir);

                // Create DataSpaceMap stream
                CreateDataSpaceMapStream(dataSpacesDir);

                // Create DataSpaceInfo directory and stream
                var dataSpaceInfoDir = dataSpacesDir.CreateDirectory("DataSpaceInfo");
                CreateDataSpaceInfoStream(dataSpaceInfoDir);

                // Create TransformInfo directory structure
                var transformInfoDir = dataSpacesDir.CreateDirectory("TransformInfo");
                var strongEncDir = transformInfoDir.CreateDirectory("StrongEncryptionTransform");
                CreatePrimaryStream(strongEncDir);
            }

            private static void CreateVersionStream(DirectoryEntry dir)
            {
                using var ms = new MemoryStream();
                using (var bw = new BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    WriteUnicodeLpp4(bw, "Microsoft.Container.DataSpaces");
                    bw.Write((ushort)1);
                    bw.Write((ushort)0);
                    bw.Write((ushort)1);
                    bw.Write((ushort)0);
                    bw.Write((ushort)1);
                    bw.Write((ushort)0);
                }

                ms.Position = 0;
                dir.CreateDocument("Version", ms);
            }

            private static void CreateDataSpaceMapStream(DirectoryEntry dir)
            {
                using var ms = new MemoryStream();
                using (var bw = new BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    bw.Write((uint)8);
                    bw.Write((uint)1);

                    long sizePos = ms.Position;
                    bw.Write((uint)0); // placeholder for size

                    bw.Write((uint)1);
                    bw.Write((uint)0);
                    WriteUnicodeLpp4(bw, "EncryptedPackage");
                    WriteUnicodeLpp4(bw, "StrongEncryptionDataSpace");

                    long endPos = ms.Position;
                    ms.Seek(sizePos, SeekOrigin.Begin);
                    bw.Write((uint)(endPos - sizePos));
                    ms.Seek(endPos, SeekOrigin.Begin);
                }

                ms.Position = 0;
                dir.CreateDocument("DataSpaceMap", ms);
            }

            private static void CreateDataSpaceInfoStream(DirectoryEntry dir)
            {
                using var ms = new MemoryStream();
                using (var bw = new BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    bw.Write((uint)8);
                    bw.Write((uint)1);
                    WriteUnicodeLpp4(bw, "StrongEncryptionTransform");
                }

                ms.Position = 0;
                dir.CreateDocument("StrongEncryptionDataSpace", ms);
            }

            private static void CreatePrimaryStream(DirectoryEntry dir)
            {
                using var ms = new MemoryStream();
                using (var bw = new BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    long hdrPos = ms.Position;
                    bw.Write((uint)0); // placeholder for header size
                    bw.Write((uint)1);
                    WriteUnicodeLpp4(bw, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");

                    long hdrEndPos = ms.Position;
                    ms.Seek(hdrPos, SeekOrigin.Begin);
                    bw.Write((uint)(hdrEndPos - hdrPos));
                    ms.Seek(hdrEndPos, SeekOrigin.Begin);

                    WriteUnicodeLpp4(bw, "Microsoft.Container.EncryptionTransform");
                    bw.Write((ushort)1);
                    bw.Write((ushort)0);
                    bw.Write((ushort)1);
                    bw.Write((ushort)0);
                    bw.Write((ushort)1);
                    bw.Write((ushort)0);
                    bw.Write((uint)0);
                    bw.Write((uint)0);
                    bw.Write((uint)0);
                    bw.Write((uint)4);
                }

                ms.Position = 0;
                dir.CreateDocument("\u0006Primary", ms);
            }

            private static void WriteUnicodeLpp4(BinaryWriter bw, string s)
            {
                byte[] bytes = System.Text.Encoding.Unicode.GetBytes(s);
                bw.Write((uint)bytes.Length);
                bw.Write(bytes);
                int pad = (4 - (bytes.Length % 4)) % 4;
                for (int i = 0; i < pad; i++)
                {
                    bw.Write((byte)0);
                }
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
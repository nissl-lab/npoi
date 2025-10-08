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

/* ================================================================
 * About NPOI
 * Author: Tony Qu
 * Author's email: tonyqus (at) gmail.com
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors: Seijiro Ikehata (modeverv at gmail.com)
 *
 * ==============================================================*/


using OpenMcdf;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace NPOI.POIFS.Crypt
{
    /// <summary>
    ///     Helper class for XLSX encryption/decryption operations.
    ///     This class provides both high-level API (FromBytesToFile, FromFileToFile)
    ///     and internal helper methods used by AgileEncryptor.
    /// </summary>
    public static class XlsxEncryptor
    {
        internal const int KeySize = 128; // AES-128
        internal const int BlockSize = 16; // 16 bytes
        internal const int SaltSize = 16; // 16 bytes
        internal const int SpinCount = 100000; // Agile default
        internal const int SegmentLength = 4096; // segment package length
        internal const int HashSize = 20; // SHA1 = 20 bytes

        public static (XDocument xmlDoc, byte[] encryptionKey, byte[] keySalt, byte[] integritySalt)
            GenerateEncryptionInfo(string password)
        {
            byte[] keySalt = RandomBytes(SaltSize); // keyData.saltValue
            byte[] verifierSalt = RandomBytes(SaltSize); // p:encryptedKey.saltValue
            byte[] pwHash = HashPassword(password, verifierSalt, SpinCount);

            // check data (verifier / verifierHash）
            byte[] verifier = RandomBytes(SaltSize);
            byte[] keySpec = RandomBytes(KeySize / 8); // actual key of AES (inside of encryptedKey)
            byte[] encryptionKey = keySpec;

            // block key same of POI
            byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
            byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
            byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };

            byte[] encryptedVerifier = HashInput(pwHash, verifierSalt, kVerifierInputBlock, verifier, KeySize / 8);

            byte[] verifierHash;
#pragma warning disable CA5350
            using(SHA1 sha = SHA1.Create())
            {
                verifierHash = sha.ComputeHash(verifier);
            }
#pragma warning restore CA5350

            byte[] encryptedVerifierHash =
                HashInput(pwHash, verifierSalt, kHashedVerifierBlock, verifierHash, KeySize / 8);

            byte[] encryptedKey = HashInput(pwHash, verifierSalt, kCryptoKeyBlock, keySpec, KeySize / 8);

            // dataIntegrity: encryptedHmacKey
            byte[] integritySalt = RandomBytes(HashSize); // HMAC key（20B）
            byte[] kIntegrityKeyBlock = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
            byte[] ivKey = GenerateIv(keySalt, kIntegrityKeyBlock, BlockSize);
            byte[] hmacKeyPadded = PadBlock(integritySalt); // bound to 16 (0 pad)
            byte[] encryptedHmacKey = EncryptWithAes(hmacKeyPadded, encryptionKey, ivKey);

            // build EncryptionInfo XML
            XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
            XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

            XElement keyDataElement = new(ns + "keyData",
                new XAttribute("blockSize", BlockSize),
                new XAttribute("cipherAlgorithm", "AES"),
                new XAttribute("cipherChaining", "ChainingModeCBC"),
                new XAttribute("hashAlgorithm", "SHA1"),
                new XAttribute("hashSize", HashSize),
                new XAttribute("keyBits", KeySize),
                new XAttribute("saltSize", SaltSize),
                new XAttribute("saltValue", Convert.ToBase64String(keySalt))
            );

            XElement dataIntegrityElement = new(ns + "dataIntegrity",
                new XAttribute("encryptedHmacKey", Convert.ToBase64String(encryptedHmacKey)),
                new XAttribute("encryptedHmacValue", "") // hmac is set after process
            );

            XElement encryptedKeyElement = new(p + "encryptedKey",
                new XAttribute("blockSize", BlockSize),
                new XAttribute("cipherAlgorithm", "AES"),
                new XAttribute("cipherChaining", "ChainingModeCBC"),
                new XAttribute("encryptedKeyValue", Convert.ToBase64String(encryptedKey)),
                new XAttribute("encryptedVerifierHashInput", Convert.ToBase64String(encryptedVerifier)),
                new XAttribute("encryptedVerifierHashValue", Convert.ToBase64String(encryptedVerifierHash)),
                new XAttribute("hashAlgorithm", "SHA1"),
                new XAttribute("hashSize", HashSize),
                new XAttribute("keyBits", KeySize),
                new XAttribute("saltSize", SaltSize),
                new XAttribute("saltValue", Convert.ToBase64String(verifierSalt)),
                new XAttribute("spinCount", SpinCount)
            );

            XDocument xmlDoc = new(
                new XElement(ns + "encryption",
                    new XAttribute(XNamespace.Xmlns + "p", p.NamespaceName),
                    keyDataElement,
                    dataIntegrityElement,
                    new XElement(ns + "keyEncryptors",
                        new XElement(ns + "keyEncryptor",
                            new XAttribute("uri", p.NamespaceName),
                            encryptedKeyElement
                        )
                    )
                )
            );

            return (xmlDoc, encryptionKey, keySalt, integritySalt);
        }

        public static void UpdateIntegrityHmac(byte[] encryptedPackage, int oleStreamSize, byte[] encryptionKey,
            byte[] keySalt, byte[] integritySalt, XDocument xmlDoc)
        {
#pragma warning disable CA5350
            using HMACSHA1 hmac = new(integritySalt);
#pragma warning restore CA5350
            // provide to hmac a beginning StreamSize(8B, little-endian)
            byte[] sizeBytes = BitConverter.GetBytes((long) oleStreamSize);
            hmac.TransformBlock(sizeBytes, 0, 8, null, 0);

            // EncryptedPackage body（exclude 8B）
            byte[] body = new byte[encryptedPackage.Length - 8];
            Buffer.BlockCopy(encryptedPackage, 8, body, 0, body.Length);
            hmac.TransformFinalBlock(body, 0, body.Length);

            // padding HMAC to 16 byte and 0 padding -> AES-CBC
            byte[] hmacValPadded = PadBlock(hmac.Hash);
            byte[] kIntegrityValueBlock = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };
            byte[] ivVal = GenerateIv(keySalt, kIntegrityValueBlock, BlockSize);
            byte[] encryptedHmacValue = EncryptWithAes(hmacValPadded, encryptionKey, ivVal);

            XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
            if(xmlDoc.Root != null)
            {
                xmlDoc.Root.Element(ns + "dataIntegrity")
                    ?.SetAttributeValue("encryptedHmacValue", Convert.ToBase64String(encryptedHmacValue));
            }
        }

        internal static byte[] CreateEncryptedBytes(XDocument encryptionInfo, byte[] encryptedData)
        {
            string tempFile = Path.GetRandomFileName();

            try
            {
                using(RootStorage root = RootStorage.Create(tempFile))
                {
                    using(CfbStream s = root.CreateStream("EncryptedPackage"))
                    {
                        s.Write(encryptedData, 0, encryptedData.Length);
                    }

                    CreateDataSpacesStructure(root);

                    using(CfbStream s2 = root.CreateStream("EncryptionInfo"))
                    using(BinaryWriter bw = new(s2))
                    {
                        bw.Write((ushort) 4);
                        bw.Write((ushort) 4);
                        bw.Write((uint) 0x40);
                        if(encryptionInfo.Root != null)
                        {
                            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                         encryptionInfo.Root.ToString(SaveOptions.DisableFormatting);
                            xml = xml.Replace(" />", "/>");
                            bw.Write(Encoding.UTF8.GetBytes(xml));
                        }
                    }

                    root.Flush();
                }

                byte[] bytes = File.ReadAllBytes(tempFile);
                return bytes;
            }
            finally
            {
                try
                {
                    if(File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }
                catch
                {
                    // no-op
                }
            }
        }

        public static void CreateEncryptedFile(Stream outputStream, XDocument encryptionInfo, byte[] encryptedData)
        {
            using MemoryStream ms = new();
            using(RootStorage root = RootStorage.Create(ms))
            {
                using(CfbStream s = root.CreateStream("EncryptedPackage"))
                {
                    s.Write(encryptedData, 0, encryptedData.Length);
                }

                CreateDataSpacesStructure(root);
                using(CfbStream s2 = root.CreateStream("EncryptionInfo"))
                using(BinaryWriter bw = new(s2))
                {
                    bw.Write((ushort) 4);
                    bw.Write((ushort) 4);
                    bw.Write((uint) 0x40);
                    if(encryptionInfo.Root != null)
                    {
                        string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                     encryptionInfo.Root.ToString(SaveOptions.DisableFormatting);
                        xml = xml.Replace(" />", "/>");
                        bw.Write(Encoding.UTF8.GetBytes(xml));
                    }
                }
            }

            ms.Position = 0;
            ms.CopyTo(outputStream);
        }

        public static byte[] EncryptPackage(byte[] packageData, byte[] encryptionKey, byte[] keySalt)
        {
            using MemoryStream ms = new();
            using BinaryWriter bw = new(ms);
            bw.Write((long) packageData.Length);
            int offset = 0;
            int block = 0;
            while(offset < packageData.Length)
            {
                int segSize = Math.Min(SegmentLength, packageData.Length - offset);
                bool isLast = offset + segSize >= packageData.Length;
                byte[] seg = new byte[segSize];
                Array.Copy(packageData, offset, seg, 0, segSize);
                byte[] blockKey = BitConverter.GetBytes(block);
                byte[] iv = GenerateIv(keySalt, blockKey, BlockSize);
                using(Aes aes = Aes.Create())
                {
                    aes.Key = encryptionKey;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
                    if(!isLast && segSize < SegmentLength)
                    {
                        byte[] padSeg = new byte[SegmentLength];
                        Array.Copy(seg, padSeg, segSize);
                        seg = padSeg;
                    }

                    using(ICryptoTransform enc = aes.CreateEncryptor())
                    {
                        bw.Write(enc.TransformFinalBlock(seg, 0, seg.Length));
                    }
                }

                offset += segSize;
                block++;
            }

            return ms.ToArray();
        }

        public static byte[] RandomBytes(int n)
        {
            byte[] b = new byte[n];
            using RandomNumberGenerator rng = RandomNumberGenerator.Create();
            rng.GetBytes(b);
            return b;
        }

        internal static byte[] HashPassword(string pw, byte[] salt, int spin)
        {
            byte[] pwb = Encoding.Unicode.GetBytes(pw);
#pragma warning disable CA5350
            using SHA1 sha = SHA1.Create();
#pragma warning restore CA5350
            try
            {
                sha.TransformBlock(salt, 0, salt.Length, null, 0);
                sha.TransformFinalBlock(pwb, 0, pwb.Length);
                byte[] h = (byte[]) sha.Hash.Clone(); // Clone to avoid issues

                for(int i = 0; i < spin; i++)
                {
                    sha.Initialize();
                    byte[] iter = BitConverter.GetBytes(i);
                    sha.TransformBlock(iter, 0, 4, null, 0);
                    sha.TransformFinalBlock(h, 0, h.Length);
                    h = (byte[]) sha.Hash.Clone();
                }

                return h;
            }
            finally
            {
                Array.Clear(pwb, 0, pwb.Length);
            }
        }

        internal static byte[] HashInput(byte[] pwHash, byte[] salt, byte[] blk, byte[] input, int keySize)
        {
            byte[] k = GenerateKey(pwHash, blk, keySize);
            byte[] iv = GenerateIv(salt, null, BlockSize);
            byte[] pad = PadBlock(input);
            return EncryptWithAes(pad, k, iv);
        }

        internal static byte[] GenerateKey(byte[] h, byte[] blk, int ks)
        {
#pragma warning disable CA5350
            using SHA1 sha = SHA1.Create();
#pragma warning restore CA5350
            sha.TransformBlock(h, 0, h.Length, null, 0);
            sha.TransformFinalBlock(blk, 0, blk.Length);
            byte[] d = sha.Hash;
            byte[] k = new byte[ks];
            Array.Copy(d, k, Math.Min(d.Length, ks));
            return k;
        }

        internal static byte[] GenerateIv(byte[] salt, byte[] blk, int bs)
        {
            if(blk == null)
            {
                byte[] iv1 = new byte[bs];
                Array.Copy(salt, iv1, Math.Min(salt.Length, bs));
                return iv1;
            }
#pragma warning disable CA5350
            using SHA1 sha = SHA1.Create();
#pragma warning restore CA5350

            sha.TransformBlock(salt, 0, salt.Length, null, 0);
            sha.TransformFinalBlock(blk, 0, blk.Length);
            byte[] d = sha.Hash;
            byte[] iv = new byte[bs];
            Array.Copy(d, iv, Math.Min(d.Length, bs));
            return iv;
        }

        internal static byte[] EncryptWithAes(byte[] d, byte[] k, byte[] iv)
        {
            using Aes aes = Aes.Create();
            aes.Key = k;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            return aes.CreateEncryptor().TransformFinalBlock(d, 0, d.Length);
        }

        internal static byte[] PadBlock(byte[] d)
        {
            int s = PadLen(d.Length);
            byte[] r = new byte[s];
            Buffer.BlockCopy(d, 0, r, 0, d.Length);
            return r;
        }

        internal static int PadLen(int len)
        {
            return (len + 15) / 16 * 16;
        }

        #region Public API (for backward compatibility)

        public static byte[] FromBytesToBytes(byte[] wbByte, string password)
        {
            ValidateEncryptionParameters();

            if(wbByte == null || wbByte.Length == 0)
            {
                throw new ArgumentException("Input data cannot be null or empty", nameof(wbByte));
            }

            if(string.IsNullOrEmpty(password))
            {
                throw new ArgumentException("Password cannot be null or empty", nameof(password));
            }

            if(password.Length > 255)
            {
                throw new ArgumentException("Password is too long (max 255 characters)", nameof(password));
            }

            try
            {
                var (xmlDoc, encryptionKey, keySalt, integritySalt) = GenerateEncryptionInfo(password);
                byte[] encryptedPackage = EncryptPackage(wbByte, encryptionKey, keySalt);
                UpdateIntegrityHmac(encryptedPackage, wbByte.Length, encryptionKey, keySalt, integritySalt, xmlDoc);
                return CreateEncryptedBytes(xmlDoc, encryptedPackage);
            }
            catch(Exception ex)
            {
                throw new InvalidOperationException("Failed to encrypt file", ex);
            }
        }

        /// <summary>
        ///     Encrypts workbook bytes and writes to a file with password protection.
        /// </summary>
        public static void FromBytesToFile(byte[] wbByte, string outputPath, string password)
        {
            byte[] bytes = FromBytesToBytes(wbByte, password);
            File.WriteAllBytes(outputPath, bytes);
        }

        /// <summary>
        ///     Encrypts an XLSX file with password protection.
        /// </summary>
        public static void FromFileToFile(string inputPath, string outputPath, string password)
        {
            byte[] packageData = File.ReadAllBytes(inputPath);
            FromBytesToFile(packageData, outputPath, password);
        }

        /// <summary>
        ///     Decrypts a password-protected XLSX file.
        /// </summary>
        public static byte[] Decrypt(string encryptedPath, string password)
        {
            if(!File.Exists(encryptedPath))
            {
                throw new FileNotFoundException("Encrypted file not found", encryptedPath);
            }

            if(string.IsNullOrEmpty(password))
            {
                throw new ArgumentException("Password cannot be null or empty", nameof(password));
            }

            using RootStorage root = RootStorage.OpenRead(encryptedPath);

            // read EncryptionInfo
            CfbStream encInfoStream;
            try
            {
                encInfoStream = root.OpenStream("EncryptionInfo");
            }
            catch
            {
                throw new InvalidOperationException("File is not encrypted (EncryptionInfo missing)");
            }

            using(encInfoStream)
            using(BinaryReader reader = new(encInfoStream))
            {
                // read version info and flag
                ushort versionMajor = reader.ReadUInt16();
                ushort versionMinor = reader.ReadUInt16();
                reader.ReadUInt32();

                if(versionMajor != 4 || versionMinor != 4)
                {
                    throw new NotSupportedException($"Unsupported encryption version: {versionMajor}.{versionMinor}");
                }

                // read XML
                byte[] xmlBytes = reader.ReadBytes((int) encInfoStream.Length - 8);
                string xmlString = Encoding.UTF8.GetString(xmlBytes);

                Match keySaltMatch = Regex.Match(xmlString, @"<keyData[^>]*saltValue=""([^""]+)""");
                Match verifierSaltMatch = Regex.Match(xmlString, @"<p:encryptedKey[^>]*saltValue=""([^""]+)""");
                Match spinCountMatch = Regex.Match(xmlString, @"spinCount=""(\d+)""");
                Match encryptedKeyMatch = Regex.Match(xmlString, @"encryptedKeyValue=""([^""]+)""");

                if(!keySaltMatch.Success || !verifierSaltMatch.Success || !spinCountMatch.Success ||
                   !encryptedKeyMatch.Success)
                {
                    throw new InvalidOperationException("fail: check encrypted info");
                }

                byte[] keySalt = Convert.FromBase64String(keySaltMatch.Groups[1].Value);
                byte[] verifierSalt = Convert.FromBase64String(verifierSaltMatch.Groups[1].Value);
                int spinCount = int.Parse(spinCountMatch.Groups[1].Value);
                byte[] encryptedKey = Convert.FromBase64String(encryptedKeyMatch.Groups[1].Value);

                if(!VerifyPassword(password, xmlString))
                {
                    throw new UnauthorizedAccessException("Invalid password");
                }

                byte[] pwHash = HashPassword(password, verifierSalt, spinCount);
                byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };
                byte[] keyIntermedKey = GenerateKey(pwHash, kCryptoKeyBlock, KeySize / 8);
                byte[] keyIv = GenerateIv(verifierSalt, null, BlockSize);

                byte[] actualKey;
                using(Aes aes = Aes.Create())
                {
                    aes.Key = keyIntermedKey;
                    aes.IV = keyIv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.None;
                    using ICryptoTransform dec = aes.CreateDecryptor();
                    byte[] decryptionKey = dec.TransformFinalBlock(encryptedKey, 0, encryptedKey.Length);
                    actualKey = new byte[KeySize / 8];
                    Array.Copy(decryptionKey, actualKey, actualKey.Length);
                }

                // read EncryptedPackage
                CfbStream encPackageStream;
                try
                {
                    encPackageStream = root.OpenStream("EncryptedPackage");
                }
                catch
                {
                    throw new InvalidOperationException("EncryptedPackage stream not found");
                }

                byte[] encryptedPackageData;
                using(encPackageStream)
                {
                    encryptedPackageData = new byte[encPackageStream.Length];
                    _ = encPackageStream.Read(encryptedPackageData, 0, encryptedPackageData.Length);
                }

                // decrypt EncryptedPackage
                byte[] decryptedData;
                long streamSize;

                using(MemoryStream ms = new(encryptedPackageData))
                using(BinaryReader br = new(ms))
                {
                    streamSize = br.ReadInt64();

                    using MemoryStream outMs = new();
                    int block = 0;
                    long remaining = streamSize;

                    while(remaining > 0)
                    {
                        int segSize = (int) Math.Min(SegmentLength, remaining);
                        bool isLast = remaining <= SegmentLength;

                        byte[] encryptedSeg = isLast
                            ? br.ReadBytes(PadLen((int) remaining))
                            : br.ReadBytes(SegmentLength);

                        byte[] blockKey = BitConverter.GetBytes(block);
                        byte[] segIv = GenerateIv(keySalt, blockKey, BlockSize);

                        using(Aes aes = Aes.Create())
                        {
                            aes.Key = actualKey;
                            aes.IV = segIv;
                            aes.Mode = CipherMode.CBC;
                            aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
                            using ICryptoTransform dec = aes.CreateDecryptor();
                            byte[] decSeg = dec.TransformFinalBlock(encryptedSeg, 0, encryptedSeg.Length);
                            outMs.Write(decSeg, 0, Math.Min(segSize, decSeg.Length));
                        }

                        remaining -= segSize;
                        block++;
                    }

                    decryptedData = outMs.ToArray();
                }

                if(!VerifyIntegrity(encryptedPackageData, (int) streamSize, actualKey, keySalt, xmlString))
                {
                    throw new InvalidOperationException(
                        "Data integrity check failed - file may be corrupted or tampered");
                }

                return decryptedData;
            }
        }

        #endregion

        #region Private Helper Methods

        private static void CreateDataSpacesStructure(RootStorage root)
        {
            OpenMcdf.Storage ds = root.CreateStorage("\u0006DataSpaces");
            using(CfbStream v = ds.CreateStream("Version"))
            using(BinaryWriter bw = new(v))
            {
                WriteUnicodeLpp4(bw, "Microsoft.Container.DataSpaces");
                bw.Write((ushort) 1);
                bw.Write((ushort) 0);
                bw.Write((ushort) 1);
                bw.Write((ushort) 0);
                bw.Write((ushort) 1);
                bw.Write((ushort) 0);
            }

            using(CfbStream m = ds.CreateStream("DataSpaceMap"))
            using(BinaryWriter bw = new(m))
            {
                bw.Write((uint) 8);
                bw.Write((uint) 1);
                long pos = m.Position;
                bw.Write((uint) 0);
                bw.Write((uint) 1);
                bw.Write((uint) 0);
                WriteUnicodeLpp4(bw, "EncryptedPackage");
                WriteUnicodeLpp4(bw, "StrongEncryptionDataSpace");
                long end = m.Position;
                m.Seek(pos, SeekOrigin.Begin);
                bw.Write((uint) (end - pos));
                m.Seek(end, SeekOrigin.Begin);
            }

            OpenMcdf.Storage dsi = ds.CreateStorage("DataSpaceInfo");
            using(CfbStream s = dsi.CreateStream("StrongEncryptionDataSpace"))
            using(BinaryWriter bw = new(s))
            {
                bw.Write((uint) 8);
                bw.Write((uint) 1);
                WriteUnicodeLpp4(bw, "StrongEncryptionTransform");
            }

            OpenMcdf.Storage ti = ds.CreateStorage("TransformInfo");
            OpenMcdf.Storage st = ti.CreateStorage("StrongEncryptionTransform");
            using(CfbStream p = st.CreateStream("\u0006Primary"))
            using(BinaryWriter bw = new(p))
            {
                long hdr = p.Position;
                bw.Write((uint) 0);
                bw.Write((uint) 1);
                WriteUnicodeLpp4(bw, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
                long hdrEnd = p.Position;
                p.Seek(hdr, SeekOrigin.Begin);
                bw.Write((uint) (hdrEnd - hdr));
                p.Seek(hdrEnd, SeekOrigin.Begin);
                WriteUnicodeLpp4(bw, "Microsoft.Container.EncryptionTransform");
                bw.Write((ushort) 1);
                bw.Write((ushort) 0);
                bw.Write((ushort) 1);
                bw.Write((ushort) 0);
                bw.Write((ushort) 1);
                bw.Write((ushort) 0);
                bw.Write((uint) 0);
                bw.Write((uint) 0);
                bw.Write((uint) 0);
                bw.Write((uint) 4);
            }
        }

        private static void WriteUnicodeLpp4(BinaryWriter bw, string s)
        {
            byte[] b = Encoding.Unicode.GetBytes(s);
            bw.Write((uint) b.Length);
            bw.Write(b);
            int pad = (4 - (b.Length % 4)) % 4;
            for(int i = 0; i < pad; i++)
            {
                bw.Write((byte) 0);
            }
        }

        private static void ValidateEncryptionParameters()
        {
            if(KeySize != 128 && KeySize != 192 && KeySize != 256)
#pragma warning disable CS0162
            {
                throw new InvalidOperationException($"Invalid key size: {KeySize}");
            }
#pragma warning restore CS0162

            if(BlockSize != 16)
#pragma warning disable CS0162
            {
                throw new InvalidOperationException($"Invalid block size: {BlockSize}");
            }
#pragma warning restore CS0162

            if(SpinCount < 1)
#pragma warning disable CS0162
            {
                throw new InvalidOperationException($"Invalid spin count: {SpinCount}");
            }
#pragma warning restore CS0162
        }

        private static bool VerifyPassword(string password, string xmlString)
        {
            Match encVerifierMatch = Regex.Match(xmlString, @"encryptedVerifierHashInput=""([^""]+)""");
            Match encVerifierHashMatch = Regex.Match(xmlString, @"encryptedVerifierHashValue=""([^""]+)""");
            Match verifierSaltMatch = Regex.Match(xmlString, @"<p:encryptedKey[^>]*saltValue=""([^""]+)""");
            Match spinCountMatch = Regex.Match(xmlString, @"spinCount=""(\d+)""");

            if(!encVerifierMatch.Success || !encVerifierHashMatch.Success)
            {
                return false;
            }

            byte[] encryptedVerifier = Convert.FromBase64String(encVerifierMatch.Groups[1].Value);
            byte[] encryptedVerifierHash = Convert.FromBase64String(encVerifierHashMatch.Groups[1].Value);
            byte[] verifierSalt = Convert.FromBase64String(verifierSaltMatch.Groups[1].Value);
            int spinCount = int.Parse(spinCountMatch.Groups[1].Value);

            byte[] pwHash = HashPassword(password, verifierSalt, spinCount);

            byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
            byte[] intermedKey = GenerateKey(pwHash, kVerifierInputBlock, KeySize / 8);
            byte[] iv = GenerateIv(verifierSalt, null, BlockSize);

            byte[] decryptedVerifier;
            using(Aes aes = Aes.Create())
            {
                aes.Key = intermedKey;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                using ICryptoTransform dec = aes.CreateDecryptor();
                decryptedVerifier = dec.TransformFinalBlock(encryptedVerifier, 0, encryptedVerifier.Length);
            }

            byte[] verifierHash;
#pragma warning disable CA5350
            using(SHA1 sha = SHA1.Create())
            {
                verifierHash = sha.ComputeHash(decryptedVerifier, 0, SaltSize);
            }
#pragma warning restore CA5350
            byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
            intermedKey = GenerateKey(pwHash, kHashedVerifierBlock, KeySize / 8);

            byte[] decryptedVerifierHash;
            using(Aes aes = Aes.Create())
            {
                aes.Key = intermedKey;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                using ICryptoTransform dec = aes.CreateDecryptor();
                decryptedVerifierHash = dec.TransformFinalBlock(encryptedVerifierHash, 0, encryptedVerifierHash.Length);
            }

            return verifierHash.Take(HashSize).SequenceEqual(decryptedVerifierHash.Take(HashSize));
        }

        private static bool VerifyIntegrity(byte[] encryptedPackage, int oleStreamSize,
            byte[] encryptionKey, byte[] keySalt, string xmlString)
        {
            Match encHmacKeyMatch = Regex.Match(xmlString, @"encryptedHmacKey=""([^""]+)""");
            Match encHmacValueMatch = Regex.Match(xmlString, @"encryptedHmacValue=""([^""]+)""");

            if(!encHmacKeyMatch.Success || !encHmacValueMatch.Success)
            {
                return false;
            }

            byte[] encryptedHmacKey = Convert.FromBase64String(encHmacKeyMatch.Groups[1].Value);
            byte[] encryptedHmacValue = Convert.FromBase64String(encHmacValueMatch.Groups[1].Value);

            // Decrypt HMAC key
            byte[] kIntegrityKeyBlock = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
            byte[] ivKey = GenerateIv(keySalt, kIntegrityKeyBlock, BlockSize);
            byte[] hmacKey;
            using(Aes aes = Aes.Create())
            {
                aes.Key = encryptionKey;
                aes.IV = ivKey;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                hmacKey = aes.CreateDecryptor().TransformFinalBlock(encryptedHmacKey, 0, encryptedHmacKey.Length);
            }

            hmacKey = hmacKey.Take(HashSize).ToArray();

            // Calculate HMAC
#pragma warning disable CA5350
            using HMACSHA1 hmac = new(hmacKey);
#pragma warning restore CA5350
            byte[] sizeBytes = BitConverter.GetBytes((long) oleStreamSize);
            hmac.TransformBlock(sizeBytes, 0, 8, null, 0);
            byte[] body = new byte[encryptedPackage.Length - 8];
            Buffer.BlockCopy(encryptedPackage, 8, body, 0, body.Length);
            hmac.TransformFinalBlock(body, 0, body.Length);

            // Decrypt expected HMAC value
            byte[] kIntegrityValueBlock = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };
            byte[] ivVal = GenerateIv(keySalt, kIntegrityValueBlock, BlockSize);
            byte[] expectedHmac;
            using(Aes aes = Aes.Create())
            {
                aes.Key = encryptionKey;
                aes.IV = ivVal;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                expectedHmac = aes.CreateDecryptor()
                    .TransformFinalBlock(encryptedHmacValue, 0, encryptedHmacValue.Length);
            }

            expectedHmac = expectedHmac.Take(HashSize).ToArray();

            return hmac.Hash.SequenceEqual(expectedHmac);
        }

        #endregion
    }
}
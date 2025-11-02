using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OpenMcdf;

namespace NPOI.POIFS.Crypt;

/// <summary>
/// decrypt
/// </summary>
public partial class XlsxEncryptor
{
    /// <summary>
    /// decrypt encrypted Excel file and return byte array
    /// </summary>
    /// <param name="encryptedPath">encrypt file path</param>
    /// <param name="password">password</param>
    /// <returns>decrypted data</returns>
    /// <exception cref="ArgumentException">pass is empty</exception>
    /// <exception cref="FileNotFoundException">file not found</exception>
    /// <exception cref="UnauthorizedAccessException">wrong password</exception>
    /// <exception cref="InvalidOperationException">decrypt</exception>
    public static byte[] Decrypt(string encryptedPath, string password)
    {
        if (string.IsNullOrEmpty(encryptedPath))
            throw new ArgumentException("Encrypted file path cannot be null or empty", nameof(encryptedPath));

        if (!File.Exists(encryptedPath))
            throw new FileNotFoundException("Encrypted file not found", encryptedPath);

        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        try
        {
            using var root = RootStorage.OpenRead(encryptedPath);
            
            var encryptionInfo = ReadEncryptionInfo(root);
            var secretKey = VerifyPasswordAndGetKey(password, encryptionInfo);
            var decryptedData = DecryptPackage(root, secretKey, encryptionInfo);
            
            return decryptedData;
        }
        catch (Exception ex) when (ex is not UnauthorizedAccessException)
        {
            throw new InvalidOperationException("Failed to decrypt file", ex);
        }
    }

    /// <summary>
    /// decrypt to file
    /// </summary>
    /// <param name="encryptedPath">encrypted file path</param>
    /// <param name="outputPath">decrypted file path</param>
    /// <param name="password">password</param>
    /// <exception cref="ArgumentException">no pass</exception>
    /// <exception cref="FileNotFoundException">file is not found</exception>
    /// <exception cref="UnauthorizedAccessException">wrong password</exception>
    /// <exception cref="InvalidOperationException">fails</exception>
    public static void DecryptToFile(string encryptedPath, string outputPath, string password)
    {
        var decryptedData = Decrypt(encryptedPath, password);
        File.WriteAllBytes(outputPath, decryptedData);
    }

    #region Private Methods

    private static EncryptionInfo ReadEncryptionInfo(RootStorage root)
    {
        CfbStream encInfoStream;
        try
        {
            encInfoStream = root.OpenStream("EncryptionInfo");
        }
        catch
        {
            throw new InvalidOperationException("File is not encrypted (EncryptionInfo missing)");
        }

        using (encInfoStream)
        using (var reader = new BinaryReader(encInfoStream))
        {
            var versionMajor = reader.ReadUInt16();
            var versionMinor = reader.ReadUInt16();
            reader.ReadUInt32(); // flags
            
            if (versionMajor != 4 || versionMinor != 4)
                throw new NotSupportedException($"Unsupported encryption version: {versionMajor}.{versionMinor}");
            
            var xmlBytes = reader.ReadBytes((int)(encInfoStream.Length - 8));
            var xmlString = Encoding.UTF8.GetString(xmlBytes);
            var xmlDoc = XDocument.Parse(xmlString);
            
            return ParseEncryptionInfo(xmlDoc);
        }
    }

    private static EncryptionInfo ParseEncryptionInfo(XDocument xmlDoc)
    {
        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
        
        var root = xmlDoc.Root;
        if (root == null)
            throw new InvalidOperationException("Invalid EncryptionInfo XML");
        
        var keyData = root.Element(ns + "keyData");
        if (keyData == null)
            throw new InvalidOperationException("keyData element not found");
        
        var info = new EncryptionInfo
        {
            BlockSize = int.Parse(keyData.Attribute("blockSize")?.Value ?? "16"),
            KeyBits = int.Parse(keyData.Attribute("keyBits")?.Value ?? "128"),
            HashAlgorithm = keyData.Attribute("hashAlgorithm")?.Value ?? "SHA1",
            HashSize = int.Parse(keyData.Attribute("hashSize")?.Value ?? "20"),
            SaltSize = int.Parse(keyData.Attribute("saltSize")?.Value ?? "16"),
            KeySalt = Convert.FromBase64String(keyData.Attribute("saltValue")?.Value ?? "")
        };
        
        var dataIntegrity = root.Element(ns + "dataIntegrity");
        if (dataIntegrity != null)
        {
            info.EncryptedHmacKey = Convert.FromBase64String(
                dataIntegrity.Attribute("encryptedHmacKey")?.Value ?? "");
            info.EncryptedHmacValue = Convert.FromBase64String(
                dataIntegrity.Attribute("encryptedHmacValue")?.Value ?? "");
        }
        
        var keyEncryptors = root.Element(ns + "keyEncryptors");
        var keyEncryptor = keyEncryptors?.Element(ns + "keyEncryptor");
        var encryptedKey = keyEncryptor?.Element(p + "encryptedKey");
        
        if (encryptedKey == null)
            throw new InvalidOperationException("encryptedKey element not found");
        
        info.VerifierSalt = Convert.FromBase64String(
            encryptedKey.Attribute("saltValue")?.Value ?? "");
        info.SpinCount = int.Parse(
            encryptedKey.Attribute("spinCount")?.Value ?? "100000");
        info.EncryptedVerifier = Convert.FromBase64String(
            encryptedKey.Attribute("encryptedVerifierHashInput")?.Value ?? "");
        info.EncryptedVerifierHash = Convert.FromBase64String(
            encryptedKey.Attribute("encryptedVerifierHashValue")?.Value ?? "");
        info.EncryptedKey = Convert.FromBase64String(
            encryptedKey.Attribute("encryptedKeyValue")?.Value ?? "");
        
        return info;
    }

    private static byte[] VerifyPasswordAndGetKey(string password, EncryptionInfo info)
    {
        using var hashAlg = CreateHashAlgorithm(info.HashAlgorithm);
        
        var pwHash = HashPassword(password, info.VerifierSalt, info.SpinCount, hashAlg);
        
        byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
        byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
        byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };
        
        var intermedKey = GenerateKey(pwHash, kVerifierInputBlock, info.KeyBits / 8, hashAlg);
        var iv = GenerateIv(info.VerifierSalt, null, info.BlockSize, hashAlg);
        
        byte[] verifierInputDec;
        using (var aes = Aes.Create())
        {
            aes.Key = intermedKey;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            using var dec = aes.CreateDecryptor();
            verifierInputDec = dec.TransformFinalBlock(info.EncryptedVerifier, 0, info.EncryptedVerifier.Length);
        }
        
        verifierInputDec = GetBlock0(verifierInputDec, info.SaltSize);
        
        byte[] verifierHash;
        using (var hash = CreateHashAlgorithm(info.HashAlgorithm))
        {
            verifierHash = hash.ComputeHash(verifierInputDec);
        }
        
        intermedKey = GenerateKey(pwHash, kHashedVerifierBlock, info.KeyBits / 8, hashAlg);
        
        byte[] verifierHashDec;
        using (var aes = Aes.Create())
        {
            aes.Key = intermedKey;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            using var dec = aes.CreateDecryptor();
            verifierHashDec = dec.TransformFinalBlock(info.EncryptedVerifierHash, 0, info.EncryptedVerifierHash.Length);
        }
        
        verifierHashDec = GetBlock0(verifierHashDec, info.HashSize);
        
        if (!verifierHash.Take(info.HashSize).SequenceEqual(verifierHashDec.Take(info.HashSize)))
        {
            throw new UnauthorizedAccessException("Invalid password");
        }
        
        intermedKey = GenerateKey(pwHash, kCryptoKeyBlock, info.KeyBits / 8, hashAlg);
        
        byte[] keySpec;
        using (var aes = Aes.Create())
        {
            aes.Key = intermedKey;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            using var dec = aes.CreateDecryptor();
            keySpec = dec.TransformFinalBlock(info.EncryptedKey, 0, info.EncryptedKey.Length);
        }
        
        keySpec = GetBlock0(keySpec, info.KeyBits / 8);
        
        return keySpec;
    }

    private static byte[] DecryptPackage(RootStorage root, byte[] secretKey, EncryptionInfo info)
    {
        var encPackageStream = root.OpenStream("EncryptedPackage");
        
        using (encPackageStream)
        using (var reader = new BinaryReader(encPackageStream))
        {
            var originalSize = reader.ReadInt64();
            
            using var ms = new MemoryStream();
            uint blockIndex = 0;
            long totalDecrypted = 0;
            const int segmentLength = 4096;
            
            using var hashAlg = CreateHashAlgorithm(info.HashAlgorithm);
            
            while (totalDecrypted < originalSize)
            {
                long remaining = originalSize - totalDecrypted;
                bool isLast = (remaining <= segmentLength);
                
                int encryptedBlockSize;
                if (isLast)
                {
                    int expectedSize = (int)remaining;
                    encryptedBlockSize = ((expectedSize / info.BlockSize) + 1) * info.BlockSize;
                }
                else
                {
                    encryptedBlockSize = segmentLength;
                }
                
                var encryptedBlock = reader.ReadBytes(encryptedBlockSize);
                
                var iv = GenerateBlockIv(info.KeySalt, blockIndex, info.BlockSize, hashAlg);
                
                byte[] decryptedBlock;
                using (var aes = Aes.Create())
                {
                    aes.Key = secretKey;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
                    
                    using var dec = aes.CreateDecryptor();
                    decryptedBlock = dec.TransformFinalBlock(encryptedBlock, 0, encryptedBlock.Length);
                }
                
                int writeSize = isLast ? (int)remaining : segmentLength;
                
                ms.Write(decryptedBlock, 0, writeSize);
                totalDecrypted += writeSize;
                blockIndex++;
            }
            
            return ms.ToArray();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Security", "CA5350:Do not use weak cryptographic algorithms", Justification = "Required for compatibility with legacy Office encryption")]    
    private static System.Security.Cryptography.HashAlgorithm CreateHashAlgorithm(string algorithmName)
    {
        return algorithmName.ToUpperInvariant() switch
        {
            "MD5" => MD5.Create(),
            "SHA1" => SHA1.Create(),
            "SHA256" => SHA256.Create(),
            "SHA384" => SHA384.Create(),
            "SHA512" => SHA512.Create(),
            _ => SHA1.Create()
        };
    }

    private static byte[] HashPassword(string pw, byte[] salt, int spin, System.Security.Cryptography.HashAlgorithm hashAlg)
    {
        var pwb = Encoding.Unicode.GetBytes(pw);

        try
        {
            hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
            hashAlg.TransformFinalBlock(pwb, 0, pwb.Length);
            var h = (byte[])hashAlg.Hash!.Clone();

            for (var i = 0; i < spin; i++)
            {
                hashAlg.Initialize();
                var iter = BitConverter.GetBytes(i);
                hashAlg.TransformBlock(iter, 0, 4, null, 0);
                hashAlg.TransformFinalBlock(h, 0, h.Length);
                h = (byte[])hashAlg.Hash!.Clone();
            }

            return h;
        }
        finally
        {
            Array.Clear(pwb, 0, pwb.Length);
        }
    }

    private static byte[] GenerateKey(byte[] h, byte[] blk, int ks, System.Security.Cryptography.HashAlgorithm hashAlg)
    {
        hashAlg.Initialize();
        hashAlg.TransformBlock(h, 0, h.Length, null, 0);
        hashAlg.TransformFinalBlock(blk, 0, blk.Length);
        var d = hashAlg.Hash!;
        var k = new byte[ks];
        Array.Copy(d, k, Math.Min(d.Length, ks));
        return k;
    }

    private static byte[] GenerateIv(byte[] salt, byte[] blk, int bs, System.Security.Cryptography.HashAlgorithm hashAlg)
    {
        if (blk == null)
        {
            var iv1 = new byte[bs];
            Array.Copy(salt, iv1, Math.Min(salt.Length, bs));
            return iv1;
        }

        hashAlg.Initialize();
        hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
        hashAlg.TransformFinalBlock(blk, 0, blk.Length);
        var d = hashAlg.Hash!;
        var iv = new byte[bs];
        Array.Copy(d, iv, Math.Min(d.Length, bs));
        return iv;
    }

    private static byte[] GenerateBlockIv(byte[] salt, uint blockKey, int bs, System.Security.Cryptography.HashAlgorithm hashAlg)
    {
        var blockBytes = BitConverter.GetBytes(blockKey);
        hashAlg.Initialize();
        hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
        hashAlg.TransformFinalBlock(blockBytes, 0, 4);
        var hash = hashAlg.Hash!;
        var iv = new byte[bs];
        Array.Copy(hash, iv, Math.Min(hash.Length, bs));
        return iv;
    }

    private static byte[] GetBlock0(byte[] data, int size)
    {
        var result = new byte[size];
        Array.Copy(data, result, Math.Min(data.Length, size));
        return result;
    }

    #endregion

    #region Internal Classes

    private sealed class EncryptionInfo
    {
        public int BlockSize { get; set; }
        public int KeyBits { get; set; }
        public string HashAlgorithm { get; set; } = "";
        public int HashSize { get; set; }
        public int SaltSize { get; set; }
        public byte[] KeySalt { get; set; } = Array.Empty<byte>();
        public byte[] VerifierSalt { get; set; } = Array.Empty<byte>();
        public int SpinCount { get; set; }
        public byte[] EncryptedVerifier { get; set; } = Array.Empty<byte>();
        public byte[] EncryptedVerifierHash { get; set; } = Array.Empty<byte>();
        public byte[] EncryptedKey { get; set; } = Array.Empty<byte>();
        public byte[] EncryptedHmacKey { get; set; } = Array.Empty<byte>();
        public byte[] EncryptedHmacValue { get; set; } = Array.Empty<byte>();
    }

    #endregion
}

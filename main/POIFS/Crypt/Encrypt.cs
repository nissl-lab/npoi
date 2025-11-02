using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OpenMcdf;

namespace NPOI.POIFS.Crypt;


public partial class XlsxEncryptor
{
    public enum AesKeySize
    {
        Aes128 = 128,
        Aes192 = 192,
        Aes256 = 256
    }
    
    public enum HashAlgorithmType
    {
        Sha1 = 20,      // 20 bytes
        Sha256 = 32,    // 32 bytes
        Sha384 = 48,    // 48 bytes
        Sha512 = 64,    // 64 bytes
        Md5 = 16        // 16 bytes (非推奨だが互換性のため)
    }
    
    
    private readonly int _keySize;
    private readonly int _blockSize = 16;
    private readonly int _saltSize = 16;
    private readonly int _spinCount = 100000;
    private readonly int _segmentLength = 4096;
    private readonly HashAlgorithmType _hashAlgorithm;
    private readonly int _hashSize;
    
    private EncryptionInfo Info
    {
        get;
        set;
    } = null;

    public XlsxEncryptor(
        AesKeySize keySize = AesKeySize.Aes128,
        HashAlgorithmType hashAlgorithm = HashAlgorithmType.Sha1)
    {
        _keySize = (int)keySize;
        _hashAlgorithm = hashAlgorithm;
        _hashSize = (int)hashAlgorithm;
        ValidateParameters();
    }

    private void ValidateParameters()
    {
        if (_keySize != 128 && _keySize != 192 && _keySize != 256)
            throw new InvalidOperationException($"Invalid key size: {_keySize}");
        
        if (_hashSize != 16 && _hashSize != 20 && _hashSize != 32 && _hashSize != 48 && _hashSize != 64)
            throw new InvalidOperationException($"Invalid hash size: {_hashSize}");
    }
    
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Security", "CA5350:Do not use weak cryptographic algorithms", Justification = "Required for compatibility with legacy Office encryption")]
    private System.Security.Cryptography.HashAlgorithm CreateHashAlgorithm()
    {
        return _hashAlgorithm switch
        {
            HashAlgorithmType.Md5 => MD5.Create(),
            HashAlgorithmType.Sha1 => SHA1.Create(),
            HashAlgorithmType.Sha256 => SHA256.Create(),
            HashAlgorithmType.Sha384 => SHA384.Create(),
            HashAlgorithmType.Sha512 => SHA512.Create(),
            _ => throw new NotSupportedException($"Hash algorithm not supported: {_hashAlgorithm}")
        };
    }

    private string GetHashAlgorithmName()
    {
        return _hashAlgorithm switch
        {
            HashAlgorithmType.Md5 => "MD5",
            HashAlgorithmType.Sha1 => "SHA1",
            HashAlgorithmType.Sha256 => "SHA256",
            HashAlgorithmType.Sha384 => "SHA384",
            HashAlgorithmType.Sha512 => "SHA512",
            _ => throw new NotSupportedException($"Hash algorithm not supported: {_hashAlgorithm}")
        };
    }
    
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Security", "CA5350:Do not use weak cryptographic algorithms", Justification = "Required for compatibility with legacy Office encryption")]    
    private HMAC CreateHmac(byte[] key)
    {
        return _hashAlgorithm switch
        {
            HashAlgorithmType.Md5 => new HMACMD5(key),
            HashAlgorithmType.Sha1 => new HMACSHA1(key),
            HashAlgorithmType.Sha256 => new HMACSHA256(key),
            HashAlgorithmType.Sha384 => new HMACSHA384(key),
            HashAlgorithmType.Sha512 => new HMACSHA512(key),
            _ => throw new NotSupportedException($"HMAC algorithm not supported: {_hashAlgorithm}")
        };
    }

    public void EncryptToFile(byte[] wbByte, string outputPath, string password)
    {
        if (wbByte == null || wbByte.Length == 0)
            throw new ArgumentException("Input data cannot be null or empty", nameof(wbByte));

        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        if (password.Length > 255)
            throw new ArgumentException("Password is too long (max 255 characters)", nameof(password));

        try
        {
            var (xmlDoc, encryptionKey, keySalt, integritySalt) = GenerateEncryptionInfo(password);
            var encryptedPackage = EncryptPackage(wbByte, encryptionKey, keySalt);
            UpdateIntegrityHmac(encryptedPackage, wbByte.Length, encryptionKey, keySalt, integritySalt, xmlDoc);
            CreateEncryptedFile(outputPath, xmlDoc, encryptedPackage);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to encrypt file", ex);
        }
    }

    public void EncryptFile(string inputPath, string outputPath, string password)
    {
        var packageData = File.ReadAllBytes(inputPath);
        EncryptToFile(packageData, outputPath, password);
    }

    public (XDocument, byte[], byte[], byte[]) GenerateEncryptionInfo(string password)
    {
        var keySalt = RandomBytes(_saltSize);
        var verifierSalt = RandomBytes(_saltSize);
        var pwHash = HashPassword(password, verifierSalt, _spinCount);

        var verifier = RandomBytes(_saltSize);
        var keySpec = RandomBytes(_keySize / 8);
        var encryptionKey = keySpec;

        byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
        byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
        byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };

        var encryptedVerifier = HashInput(pwHash, verifierSalt, kVerifierInputBlock, verifier, _keySize / 8);

        byte[] verifierHash;
        using (var hashAlg = CreateHashAlgorithm())
        {
            verifierHash = hashAlg.ComputeHash(verifier);
        }

        var encryptedVerifierHash = HashInput(pwHash, verifierSalt, kHashedVerifierBlock, verifierHash, _keySize / 8);
        var encryptedKey = HashInput(pwHash, verifierSalt, kCryptoKeyBlock, keySpec, _keySize / 8);

        var integritySalt = RandomBytes(_hashSize);
        byte[] kIntegrityKeyBlock = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
        var ivKey = GenerateIv(keySalt, kIntegrityKeyBlock, _blockSize);
        var hmacKeyPadded = PadBlock(integritySalt);
        var encryptedHmacKey = EncryptWithAes(hmacKeyPadded, encryptionKey, ivKey, false);

        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        var keyDataElement = new XElement(ns + "keyData",
            new XAttribute("blockSize", _blockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("hashAlgorithm", GetHashAlgorithmName()),
            new XAttribute("hashSize", _hashSize), 
            new XAttribute("keyBits", _keySize),
            new XAttribute("saltSize", _saltSize),
            new XAttribute("saltValue", Convert.ToBase64String(keySalt))
        );

        var dataIntegrityElement = new XElement(ns + "dataIntegrity",
            new XAttribute("encryptedHmacKey", Convert.ToBase64String(encryptedHmacKey)),
            new XAttribute("encryptedHmacValue", "")
        );

        var encryptedKeyElement = new XElement(p + "encryptedKey",
            new XAttribute("blockSize", _blockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("encryptedKeyValue", Convert.ToBase64String(encryptedKey)),
            new XAttribute("encryptedVerifierHashInput", Convert.ToBase64String(encryptedVerifier)),
            new XAttribute("encryptedVerifierHashValue", Convert.ToBase64String(encryptedVerifierHash)),
            new XAttribute("hashAlgorithm", GetHashAlgorithmName()),
            new XAttribute("hashSize", _hashSize),
            new XAttribute("keyBits", _keySize),
            new XAttribute("saltSize", _saltSize),
            new XAttribute("saltValue", Convert.ToBase64String(verifierSalt)),
            new XAttribute("spinCount", _spinCount)
        );

        var xmlDoc = new XDocument(
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

    /// <summary>
    /// EncryptPackage
    /// </summary>
    private byte[] EncryptPackage(byte[] data, byte[] key, byte[] keySalt)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);
        
        writer.Write((long)data.Length);
        
        int offset = 0;
        uint blockIndex = 0;

        while (offset < data.Length)
        {
            int blockSize = Math.Min(_segmentLength, data.Length - offset);
            bool isLast = (offset + blockSize >= data.Length);
            
            var iv = GenerateBlockIv(keySalt, blockIndex, _blockSize);
            
            byte[] block;
            if (isLast)
            {
                block = new byte[blockSize];
                Buffer.BlockCopy(data, offset, block, 0, blockSize);
            }
            else
            {
                block = new byte[_segmentLength];
                Buffer.BlockCopy(data, offset, block, 0, blockSize);
            }
            
            var encrypted = EncryptWithAes(block, key, iv, isLast);
            writer.Write(encrypted);
            
            offset += blockSize;
            blockIndex++;
        }

        return ms.ToArray();
    }

    private void UpdateIntegrityHmac(byte[] encryptedPackage, int oleStreamSize, byte[] encryptionKey,
        byte[] keySalt, byte[] integritySalt, XDocument xmlDoc)
    {
        using var hmac = CreateHmac(integritySalt);
        var sizeBytes = BitConverter.GetBytes((long)oleStreamSize);
        hmac.TransformBlock(sizeBytes, 0, 8, null, 0);

        var body = new byte[encryptedPackage.Length - 8];
        Buffer.BlockCopy(encryptedPackage, 8, body, 0, body.Length);
        hmac.TransformFinalBlock(body, 0, body.Length);

        var hmacValPadded = PadBlock(hmac.Hash);
        byte[] kIntegrityValueBlock = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };
        var ivVal = GenerateIv(keySalt, kIntegrityValueBlock, _blockSize);
        var encryptedHmacValue = EncryptWithAes(hmacValPadded, encryptionKey, ivVal, false);

        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        if (xmlDoc.Root != null)
            xmlDoc.Root.Element(ns + "dataIntegrity")
                ?.SetAttributeValue("encryptedHmacValue", Convert.ToBase64String(encryptedHmacValue));
    }

    private static void CreateEncryptedFile(string outputPath, XDocument xmlDoc, byte[] encryptedPackage)
    {
        using var root = RootStorage.Create(outputPath);
        
        using (var encInfoStream = root.CreateStream("EncryptionInfo"))
        using (var writer = new BinaryWriter(encInfoStream))
        {
            writer.Write((ushort)4);
            writer.Write((ushort)4);
            writer.Write((uint)0x40);
            
            var xmlString = xmlDoc.ToString(SaveOptions.DisableFormatting);
            var xmlBytes = Encoding.UTF8.GetBytes(xmlString);
            writer.Write(xmlBytes);
        }

        using (var encPackageStream = root.CreateStream("EncryptedPackage"))
        {
            encPackageStream.Write(encryptedPackage, 0, encryptedPackage.Length);
        }
    }

    private static byte[] RandomBytes(int length)
    {
        var bytes = new byte[length];
        using var rng = RandomNumberGenerator.Create();
        rng.GetBytes(bytes);
        return bytes;
    }

    private static byte[] PadBlock(byte[] data)
    {
        int padded = ((data.Length + 15) / 16) * 16;
        var result = new byte[padded];
        Buffer.BlockCopy(data, 0, result, 0, data.Length);
        return result;
    }

    private byte[] HashPassword(string pw, byte[] salt, int spin)
    {
        var pwb = Encoding.Unicode.GetBytes(pw);
        using var hashAlg = CreateHashAlgorithm();

        try
        {
            hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
            hashAlg.TransformFinalBlock(pwb, 0, pwb.Length);
            var h = (byte[])hashAlg.Hash.Clone();

            for (var i = 0; i < spin; i++)
            {
                hashAlg.Initialize();
                var iter = BitConverter.GetBytes(i);
                hashAlg.TransformBlock(iter, 0, 4, null, 0);
                hashAlg.TransformFinalBlock(h, 0, h.Length);
                h = (byte[])hashAlg.Hash.Clone();
            }

            return h;
        }
        finally
        {
            Array.Clear(pwb, 0, pwb.Length);
        }
    }

    private byte[] HashInput(byte[] pwHash, byte[] salt, byte[] blk, byte[] input, int keySize)
    {
        var k = GenerateKey(pwHash, blk, keySize);
        var iv = GenerateIv(salt, null, _blockSize);
        var pad = PadBlock(input);
        return EncryptWithAes(pad, k, iv, false);
    }

    private byte[] GenerateKey(byte[] h, byte[] blk, int ks)
    {
        using var hashAlg = CreateHashAlgorithm();
        hashAlg.TransformBlock(h, 0, h.Length, null, 0);
        hashAlg.TransformFinalBlock(blk, 0, blk.Length);
        var d = hashAlg.Hash;
        var k = new byte[ks];
        Array.Copy(d, k, Math.Min(d.Length, ks));
        return k;
    }

    private byte[] GenerateIv(byte[] salt, byte[] blk, int bs)
    {
        if (blk == null)
        {
            var iv1 = new byte[bs];
            Array.Copy(salt, iv1, Math.Min(salt.Length, bs));
            return iv1;
        }

        using var hashAlg = CreateHashAlgorithm();
        hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
        hashAlg.TransformFinalBlock(blk, 0, blk.Length);
        var d = hashAlg.Hash;
        var iv = new byte[bs];
        Array.Copy(d, iv, Math.Min(d.Length, bs));
        return iv;
    }

    private byte[] GenerateBlockIv(byte[] salt, uint blockKey, int bs)
    {
        var blockBytes = BitConverter.GetBytes(blockKey);
        using var hashAlg = CreateHashAlgorithm();
        hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
        hashAlg.TransformFinalBlock(blockBytes, 0, 4);
        var hash = hashAlg.Hash;
        var iv = new byte[bs];
        Array.Copy(hash, iv, Math.Min(hash.Length, bs));
        return iv;
    }

    private static byte[] EncryptWithAes(byte[] d, byte[] k, byte[] iv, bool isLast)
    {
        using var aes = Aes.Create();
        aes.Key = k;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
        return aes.CreateEncryptor().TransformFinalBlock(d, 0, d.Length);
    }
    
    public static void FromBytesToFile(byte[] bytes, string outputPath, string passwordString)
    {
        var encryptor = new XlsxEncryptor();
        encryptor.EncryptToFile(bytes, outputPath, passwordString);
    }
    
    public static void FromFileToFile(string inputPath, string outputPath, string passwordString)
    {
        var encryptor = new XlsxEncryptor();
        encryptor.EncryptFile(inputPath, outputPath, passwordString);
    }
    
    public static byte[] FromBytesToBytes(
        byte[] bytes, 
        string encryptorPassword,
        AesKeySize keySize = AesKeySize.Aes128,
        HashAlgorithmType hashAlgorithm = HashAlgorithmType.Sha1
        )
    {
        var encryptor = new XlsxEncryptor(keySize,hashAlgorithm);
            
        using var ms = new MemoryStream();
        var tempFile = Path.GetRandomFileName();
            
        try
        {
            encryptor.EncryptToFile(bytes, tempFile, encryptorPassword);
            return File.ReadAllBytes(tempFile);
        }
        finally
        {
　          if (File.Exists(tempFile))
           {
               File.Delete(tempFile);
           }
        }
    }
}

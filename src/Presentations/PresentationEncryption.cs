using System;
using System.Buffers.Binary;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OpenMcdf;

namespace ShapeCrawler.Presentations;

// MS-OFFCRYPTO 2.3.4.10-15: Office Agile encryption, AES-256/CBC and SHA-512.
internal static class PresentationEncryption
{
    private const int SpinCount = 100000;
    private static readonly XNamespace EncryptionNamespace = "http://schemas.microsoft.com/office/2006/encryption";
    private static readonly XNamespace PasswordNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

    internal static void ValidatePassword(string password)
    {
        ArgumentException.ThrowIfNullOrEmpty(password);
        if (password.Length > 255)
        {
            throw new ArgumentException("Office passwords must not exceed 255 characters.", nameof(password));
        }
    }

    internal static void Save(Stream package, Stream destination, string password)
    {
        var key = RandomNumberGenerator.GetBytes(32);
        var salt = RandomNumberGenerator.GetBytes(16);
        var passwordSalt = RandomNumberGenerator.GetBytes(16);
        var passwordHash = HashPassword(password, passwordSalt);
        try
        {
            using var container = new MemoryStream();
            using (var root = RootStorage.Create(container, flags: StorageModeFlags.LeaveOpen))
            {
                var hmacKey = RandomNumberGenerator.GetBytes(64);
                byte[] hmac;
                using (var payload = root.CreateStream("EncryptedPackage"))
                {
                    WritePayload(package, payload, key, salt);
                    payload.Position = 0;
                    hmac = HMACSHA512.HashData(hmacKey, payload);
                }
                var verifier = RandomNumberGenerator.GetBytes(16);
                var encryptedKey = Parameters(PasswordNamespace + "encryptedKey", passwordSalt);
                encryptedKey.Add(
                    new XAttribute("spinCount", SpinCount),
                    Attribute("encryptedVerifierHashInput", EncryptPasswordValue(verifier, passwordHash, passwordSalt, "FEA7D2763B4B9E79")),
                    Attribute("encryptedVerifierHashValue", EncryptPasswordValue(SHA512.HashData(verifier), passwordHash, passwordSalt, "D7AA0F6D3061344E")),
                    Attribute("encryptedKeyValue", EncryptPasswordValue(key, passwordHash, passwordSalt, "146E0BE7ABACD0D6")));
                var descriptor = new XElement(
                    EncryptionNamespace + "encryption",
                    new XAttribute(XNamespace.Xmlns + "p", PasswordNamespace),
                    Parameters(EncryptionNamespace + "keyData", salt),
                    new XElement(
                        EncryptionNamespace + "dataIntegrity",
                        Attribute("encryptedHmacKey", Encrypt(hmacKey, key, HashBlock(salt, Convert.FromHexString("5FB2AD010CB9E1F6"))[..16])),
                        Attribute("encryptedHmacValue", Encrypt(hmac, key, HashBlock(salt, Convert.FromHexString("A0677F02B22C8433"))[..16]))),
                    new XElement(
                        EncryptionNamespace + "keyEncryptors",
                        new XElement(
                            EncryptionNamespace + "keyEncryptor",
                            new XAttribute("uri", PasswordNamespace.NamespaceName),
                            encryptedKey)));
                using (var info = root.CreateStream("EncryptionInfo"))
                {
                    info.Write(Convert.FromHexString("0400040040000000"));
                    info.Write(Encoding.UTF8.GetBytes(descriptor.ToString(SaveOptions.DisableFormatting)));
                }

                EncryptionDataSpaces.Write(root);
            }

            container.Position = 0;
            if (destination.CanSeek)
            {
                destination.Position = 0;
                destination.SetLength(0);
            }

            container.CopyTo(destination);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(key);
            CryptographicOperations.ZeroMemory(passwordHash);
        }
    }

    private static XAttribute Attribute(string name, byte[] value) => new(name, Convert.ToBase64String(value));

    private static XElement Parameters(XName name, byte[] salt) => new(
        name,
        new XAttribute("saltSize", 16),
        new XAttribute("blockSize", 16),
        new XAttribute("keyBits", 256),
        new XAttribute("hashSize", 64),
        new XAttribute("cipherAlgorithm", "AES"),
        new XAttribute("cipherChaining", "ChainingModeCBC"),
        new XAttribute("hashAlgorithm", "SHA512"),
        Attribute("saltValue", salt));

    private static byte[] HashPassword(string password, byte[] salt)
    {
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        try
        {
            var hash = HashBlock(salt, passwordBytes);
            Span<byte> iteration = stackalloc byte[68];
            for (var i = 0; i < SpinCount; i++)
            {
                BinaryPrimitives.WriteInt32LittleEndian(iteration, i);
                hash.CopyTo(iteration[4..]);
                SHA512.HashData(iteration, hash);
            }

            return hash;
        }
        finally
        {
            CryptographicOperations.ZeroMemory(passwordBytes);
        }
    }

    private static byte[] HashBlock(byte[] salt, byte[] block)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA512);
        hash.AppendData(salt);
        hash.AppendData(block);
        return hash.GetHashAndReset();
    }

    private static byte[] EncryptPasswordValue(byte[] value, byte[] passwordHash, byte[] salt, string block)
    {
        var derivedKey = HashBlock(passwordHash, Convert.FromHexString(block))[..32];
        try
        {
            return Encrypt(value, derivedKey, salt);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(derivedKey);
        }
    }

    private static byte[] Encrypt(byte[] value, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.Key = key;
        return aes.EncryptCbc(value, iv, PaddingMode.Zeros);
    }

    private static void WritePayload(Stream package, Stream payload, byte[] key, byte[] salt)
    {
        Span<byte> length = stackalloc byte[8];
        BinaryPrimitives.WriteInt64LittleEndian(length, package.Length);
        payload.Write(length);
        var buffer = new byte[4096];
        var block = new byte[4];
        var index = 0;
        int count;
        while ((count = package.ReadAtLeast(buffer, buffer.Length, throwOnEndOfStream: false)) > 0)
        {
            BinaryPrimitives.WriteInt32LittleEndian(block, index++);
            var iv = HashBlock(salt, block)[..16];
            payload.Write(Encrypt(buffer[..count], key, iv));
        }
    }
}

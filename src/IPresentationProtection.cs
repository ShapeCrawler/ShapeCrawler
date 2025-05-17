using System;

namespace ShapeCrawler;

/// <summary>
/// Provides methods for encrypting and decrypting PowerPoint presentations.
/// </summary>
public interface IPresentationProtection
{
    /// <summary>
    /// Gets a value indicating whether the presentation is encrypted.
    /// </summary>
    bool IsEncrypted { get; }
    
    /// <summary>
    /// Encrypts the presentation with the specified password and encryption strength.
    /// </summary>
    /// <param name="password">The password to encrypt the presentation with.</param>
    /// <param name="strength">The encryption strength (defaults to AES-256).</param>
    void Encrypt(string password, EncryptionStrength strength = EncryptionStrength.Aes256);
    
    /// <summary>
    /// Attempts to decrypt the presentation with the specified password.
    /// </summary>
    /// <param name="password">The password to decrypt the presentation with.</param>
    /// <returns><c>true</c> if decryption was successful; otherwise, <c>false</c>.</returns>
    bool TryDecrypt(string password);
}

/// <summary>
/// Specifies the encryption strength for presentation protection.
/// </summary>
public enum EncryptionStrength
{
    /// <summary>
    /// AES-128 encryption.
    /// </summary>
    Aes128,
    
    /// <summary>
    /// AES-256 encryption (stronger, default).
    /// </summary>
    Aes256
}

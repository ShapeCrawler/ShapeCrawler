using System;
using System.IO;
using System.Security.Cryptography;

namespace ShapeCrawler.ShapeCollection;

internal readonly record struct ImageStream(Stream Stream)
{
    internal string Base64Hash
    {
        get
        {
            using var sha512 = SHA512.Create();
            this.Stream.Position = 0;
            var hash = sha512.ComputeHash(this.Stream);
            this.Stream.Position = 0;

            return Convert.ToBase64String(hash);
        }
    }
}
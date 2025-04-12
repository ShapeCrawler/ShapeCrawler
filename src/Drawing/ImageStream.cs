using System;
using System.IO;
using System.Security.Cryptography;

namespace ShapeCrawler.Drawing;

internal readonly ref struct ImageStream(Stream stream)
{
    internal string Base64Hash
    {
        get
        {
            using var sha512 = SHA512.Create();
            stream.Position = 0;
            var hash = sha512.ComputeHash(stream);
            stream.Position = 0;

            return Convert.ToBase64String(hash);
        }
    }
}
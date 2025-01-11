using System;
using System.IO;
using System.Security.Cryptography;
using SkiaSharp;

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

    internal string Mime
    {
        get
        {
            var imageStreamCopy = new MemoryStream(); // the below disposes the underlying stream
            this.Stream.CopyTo(imageStreamCopy);
            imageStreamCopy.Position = 0;
            using var codec = SKCodec.Create(imageStreamCopy);
            var mime = codec.EncodedFormat switch
            {
                SKEncodedImageFormat.Jpeg => "image/jpeg",
                SKEncodedImageFormat.Png => "image/png",
                SKEncodedImageFormat.Gif => "image/gif",
                SKEncodedImageFormat.Bmp => "image/bmp",
                _ => "image/png"
            };

            return mime;
        }
    }
}
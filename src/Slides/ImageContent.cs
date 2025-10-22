using System;
using System.IO;
using System.Linq;
using ImageMagick;
using ImageMagick.Formats;
using ShapeCrawler.Drawing;

namespace ShapeCrawler.Slides;

/// <summary>
///     Represents processed image content ready for insertion into a PowerPoint slide.
/// </summary>
internal sealed class ImageContent
{
    private readonly MagickImage image;
    private readonly Stream originalStream;
    private readonly MagickFormat originalFormat;

    public ImageContent(Stream imageStream)
    {
        this.originalStream = imageStream;

        if (imageStream.CanSeek)
        {
            imageStream.Position = 0;
        }

        this.image = CreateMagickImage(imageStream);
        this.originalFormat = this.image.Format;

        EnsureSupportedImageFormat(this.image);
        HandleSvgFormat(this.image, this.originalFormat);

        var width = this.image.Width;
        var height = this.image.Height;

        if (this.originalFormat == MagickFormat.Svg)
        {
            ResizeSvgImageIfNeeded(this.image, ref width, ref height);
        }

        this.Width = width;
        this.Height = height;
    }

    public uint Width { get; private set; }

    public uint Height { get; }

    public bool IsSvg => this.originalFormat == MagickFormat.Svg;

    public bool IsOriginalFormatPreserved =>
        this.originalFormat is MagickFormat.Gif or MagickFormat.Jpeg or MagickFormat.Png or MagickFormat.Tif or MagickFormat.Tiff;

    public string MimeType => GetMimeType(this.IsOriginalFormatPreserved ? this.originalFormat : this.image.Format);

    public string Hash
    {
        get
        {
            if (this.IsOriginalFormatPreserved)
            {
                this.originalStream.Position = 0;
                return new ImageStream(this.originalStream).Base64Hash;
            }

            using var rasterStream = this.GetRasterStream();
            return new ImageStream(rasterStream).Base64Hash;
        }
    }

    public string SvgHash
    {
        get
        {
            this.originalStream.Position = 0;
            return new ImageStream(this.originalStream).Base64Hash;
        }
    }

    /// <summary>
    ///     Gets the raster stream for the image.
    /// </summary>
    public MemoryStream GetRasterStream()
    {
        var rasterStream = new MemoryStream();
        this.image.Settings.SetDefines(new PngWriteDefines { ExcludeChunks = PngChunkFlags.date });
        this.image.Settings.SetDefine("png:exclude-chunk", "tIME");
        this.image.Write(rasterStream);
        rasterStream.Position = 0;
        return rasterStream;
    }

    /// <summary>
    ///     Gets the original stream for formats that are preserved as-is.
    /// </summary>
    public Stream GetOriginalStream()
    {
        this.originalStream.Position = 0;
        return this.originalStream;
    }

    private static MagickImage CreateMagickImage(Stream imageStream)
    {
        var format = IsIco(imageStream)
            ? MagickFormat.Ico
            : MagickFormat.Unknown;

        return new MagickImage(
            imageStream,
            new MagickReadSettings
            {
                Format = format,
                BackgroundColor = MagickColors.Transparent,
            });
    }

    private static bool IsIco(Stream stream)
    {
        if (stream.Length < 6)
        {
            return false;
        }

        var originalPosition = stream.Position;
        stream.Seek(0, SeekOrigin.Begin);

        try
        {
            var header = new byte[6];
            var bytesRead = stream.Read(header, 0, 6);

            if (bytesRead < 6)
            {
                return false;
            }

            // ICO file signature:
            // Bytes 0-1: Reserved (must be 0x00 0x00)
            // Bytes 2-3: Image type (must be 0x01 0x00 for ICO)
            // Bytes 4-5: Number of images (must be > 0)
            // https://docs.fileformat.com/image/ico/#header
            return header[0] == 0x00 &&
                   header[1] == 0x00 &&
                   header[2] == 0x01 &&
                   header[3] == 0x00 &&
                   (header[4] > 0 || header[5] > 0);
        }
        finally
        {
            stream.Seek(originalPosition, SeekOrigin.Begin);
        }
    }

    private static void EnsureSupportedImageFormat(MagickImage image)
    {
        MagickFormat[] supportedImageFormats =
        [
            MagickFormat.Jpeg,
            MagickFormat.Png,
            MagickFormat.Gif,
            MagickFormat.Tif,
            MagickFormat.Tiff,
            MagickFormat.Svg
        ];

        if (!supportedImageFormats.Contains(image.Format))
        {
            image.Format = image.HasAlpha ? MagickFormat.Png : MagickFormat.Jpeg;
        }
    }

    private static void HandleSvgFormat(MagickImage image, MagickFormat originalFormat)
    {
        if (originalFormat == MagickFormat.Svg)
        {
            image.Format = MagickFormat.Png;
            image.Density =
                new Density(384, DensityUnit.PixelsPerInch); // in PowerPoint, the resolution of the rasterized version of SVG is set to 384 PPI
        }
    }

    private static void ResizeSvgImageIfNeeded(MagickImage image, ref uint width, ref uint height)
    {
        if (height > 500 || width > 500)
        {
            height = height > 500 ? 500 : height;
            width = width > 500 ? 500 : width;
            width = height == 500 ? (uint)(height * image.Width / (decimal)image.Height) : width;
            height = width == 500 ? (uint)(width * image.Height / (decimal)image.Width) : height;
            image.Resize(width, height);
        }
    }

    private static string GetMimeType(MagickFormat format)
    {
        var mime = MagickFormatInfo.Create(format)?.MimeType;

        return mime ?? throw new SCException("Unsupported image format.");
    }
}

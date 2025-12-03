using System;
using System.IO;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

internal sealed class SlideImage(ISlide slide)
{
    internal void Save(Stream stream, SKEncodedImageFormat format)
    {
        var presPart = slide.GetSDKPresentationPart();
        var pSlideSize = presPart.Presentation.SlideSize!;
        var width = pSlideSize.Cx!.Value / 9525; // Convert EMUs to pixels (96 DPI)
        var height = pSlideSize.Cy!.Value / 9525;

        using var surface = SKSurface.Create(new SKImageInfo(width, height));
        var canvas = surface.Canvas;

        this.RenderBackground(canvas);

        using var image = surface.Snapshot();
        using var data = image.Encode(format, 100);
        data.SaveTo(stream);
    }

    private static SKColor HexToSkColor(string hex)
    {
        hex = hex.TrimStart('#');

        if (hex.Length == 6)
        {
            // Parse RGB (RRGGBB)
            var r = Convert.ToByte(hex[..2], 16);
            var g = Convert.ToByte(hex.Substring(2, 2), 16);
            var b = Convert.ToByte(hex.Substring(4, 2), 16);
            return new SKColor(r, g, b);
        }

        if (hex.Length == 8)
        {
            // Parse ARGB (AARRGGBB)
            var a = Convert.ToByte(hex[..2], 16);
            var r = Convert.ToByte(hex.Substring(2, 2), 16);
            var g = Convert.ToByte(hex.Substring(4, 2), 16);
            var b = Convert.ToByte(hex.Substring(6, 2), 16);
            return new SKColor(r, g, b, a);
        }

        // Default to white if hex is invalid
        return SKColors.White;
    }

    private void RenderBackground(SKCanvas canvas)
    {
        var fill = slide.Fill;

        if (fill is { Type: FillType.Solid, Color: not null })
        {
            var skColor = HexToSkColor(fill.Color);
            canvas.Clear(skColor);
        }
        else if (fill is { Type: FillType.Picture, Picture: not null })
        {
            var bytes = fill.Picture.AsByteArray();
            using var stream = new MemoryStream(bytes);
            using var bitmap = SKBitmap.Decode(stream);
            var destRect = new SKRect(0, 0, canvas.DeviceClipBounds.Width, canvas.DeviceClipBounds.Height);
            canvas.DrawBitmap(bitmap, destRect);
        }
        else
        {
            // Default to white for unsupported backgrounds
            canvas.Clear(SKColors.White);
        }
    }
}
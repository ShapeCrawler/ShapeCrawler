using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Slides;

/// <inheritdoc/>
internal sealed class DrawingSlide(ILayoutSlide layoutSlide, UserSlideShapeCollection shapes, SlidePart slidePart)
    : UserSlide(layoutSlide, shapes, slidePart)
{
    /// <inheritdoc/>
    public override void SaveImage(string file)
    {
        using var fileStream = File.Create(file);
        this.SaveImage(fileStream);
    }

    /// <inheritdoc/>
    public override void SaveImage(Stream stream)
    {
        var presPart = ((PresentationDocument)this.GetSdkSlidePart().OpenXmlPackage).PresentationPart!;
        var pSlideSize = presPart.Presentation!.SlideSize!;
        var width = new Emus(pSlideSize.Cx!.Value).AsPixels();
        var height = new Emus(pSlideSize.Cy!.Value).AsPixels();

        this.RenderImage(stream, (int)width, (int)height, 1, 1);
    }

    /// <inheritdoc/>
    public override void SaveImage(string file, int width, int height)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

        using var fileStream = File.Create(file);
        this.SaveImage(fileStream, width, height);
    }

    /// <inheritdoc/>
    public override void SaveImage(Stream stream, int width, int height)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

        var presPart = ((PresentationDocument)this.GetSdkSlidePart().OpenXmlPackage).PresentationPart!;
        var pSlideSize = presPart.Presentation!.SlideSize!;
        var slideWidth = new Emus(pSlideSize.Cx!.Value).AsPixels();
        var slideHeight = new Emus(pSlideSize.Cy!.Value).AsPixels();

        this.RenderImage(stream, width, height, (float)(width / slideWidth), (float)(height / slideHeight));
    }

    private void RenderImage(Stream stream, int width, int height, float scaleX, float scaleY)
    {
        using var surface = SKSurface.Create(new SKImageInfo(width, height))
            ?? throw new SCException("Unable to allocate the slide image at the requested pixel dimensions.");

        var canvas = surface.Canvas;

        this.RenderBackground(canvas);
        canvas.Scale(scaleX, scaleY);
        shapes.Render(canvas);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        data.SaveTo(stream);

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }
    }

    private SKColor GetSkColor()
    {
        var hex = this.Fill.Color!.TrimStart('#');

        // Validate hex length before parsing.
        if (hex.Length != 6 && hex.Length != 8)
        {
            return SKColors.White; // used by the PowerPoint application as the default background color
        }

        return new Color(hex).AsSkColor();
    }

    private void RenderBackground(SKCanvas canvas)
    {
        var slideFill = this.Fill;
        switch (slideFill)
        {
            case { Type: FillType.Solid, Color: not null }:
                {
                    var skColor = this.GetSkColor();
                    canvas.Clear(skColor);
                    break;
                }

            case { Type: FillType.Picture, Picture: not null }:
                {
                    var bytes = slideFill.Picture.AsByteArray();
                    using var stream = new MemoryStream(bytes);
                    using var bitmap = SKBitmap.Decode(stream);
                    var destRect = new SKRect(0, 0, canvas.DeviceClipBounds.Width, canvas.DeviceClipBounds.Height);
                    canvas.DrawBitmap(bitmap, destRect, SKSamplingOptions.Default);
                    break;
                }

            default:
                // Default to white for unsupported backgrounds.
                canvas.Clear(SKColors.White);
                break;
        }
    }
}

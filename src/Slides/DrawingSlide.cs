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

        new SlideBackground(this.GetSdkSlidePart()).Render(canvas);
        canvas.Scale(scaleX, scaleY);
        new SlideTemplate(this.GetSdkSlidePart()).Render(canvas);
        shapes.Render(canvas);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        data.SaveTo(stream);

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }
    }

}

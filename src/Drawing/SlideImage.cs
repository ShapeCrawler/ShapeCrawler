using System.IO;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

internal sealed class SlideImage
{
    private readonly ISlide slide;

    internal SlideImage(ISlide slide)
    {
        this.slide = slide;
    }

    internal void Save(Stream stream, SKEncodedImageFormat format)
    {
        // TODO: Implement rendering logic
        // For now, just create a blank image with background color

        // 1. Determine slide size (default to 960x540 for now if not available)
        int width = 960;
        int height = 540;
        var presPart = this.slide.GetSDKPresentationPart();
        var pSlideSize = presPart.Presentation.SlideSize;
        if (pSlideSize?.Cx != null && pSlideSize.Cy != null)
        {
            width = pSlideSize.Cx.Value / 9525; // Convert EMUs to pixels (96 DPI)
            height = pSlideSize.Cy.Value / 9525;
        }

        using var surface = SKSurface.Create(new SKImageInfo(width, height));
        var canvas = surface.Canvas;

        // 2. Clear with white background
        canvas.Clear(SKColors.White);

        // 3. Save to stream
        using var image = surface.Snapshot();
        using var data = image.Encode(format, 100);
        data.SaveTo(stream);
    }
}
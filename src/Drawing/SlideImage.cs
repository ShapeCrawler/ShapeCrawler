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
        // TODO: Get actual slide size from Presentation
        const int width = 960;
        const int height = 540;

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

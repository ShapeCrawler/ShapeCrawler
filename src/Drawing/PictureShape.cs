using System;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class PictureShape(Picture picture, P.Picture pPicture) : Shape(new Position(pPicture),
    new ShapeSize(pPicture), new ShapeId(pPicture), pPicture)
{
    public override IPicture Picture => picture;

    public override void CopyTo(P.ShapeTree pShapeTree) => picture.CopyTo(pShapeTree);

    internal override void Render(SKCanvas canvas)
    {
        var picture = this.Picture;
        var image = picture.Image;
        if (image is null)
        {
            return;
        }

        var imageBytes = image.AsByteArray();
        using var bitmap = SKBitmap.Decode(imageBytes);
        if (bitmap is null)
        {
            return;
        }

        var x = new Points(this.X).AsPixels();
        var y = new Points(this.Y).AsPixels();
        var width = new Points(this.Width).AsPixels();
        var height = new Points(this.Height).AsPixels();

        canvas.Save();
        ApplyRotation(canvas, this, this.X, this.Y, this.Width, this.Height);

        var crop = picture.Crop;
        var srcLeft = (float)(bitmap.Width * (double)(crop.Left / 100m));
        var srcTop = (float)(bitmap.Height * (double)(crop.Top / 100m));
        var srcRight = (float)(bitmap.Width * (1 - (double)(crop.Right / 100m)));
        var srcBottom = (float)(bitmap.Height * (1 - (double)(crop.Bottom / 100m)));
        var srcRect = new SKRect(srcLeft, srcTop, srcRight, srcBottom);

        var destRect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        using var paint = new SKPaint();
        paint.IsAntialias = true;

        var transparency = picture.Transparency;
        if (transparency > 0)
        {
            var alpha = (byte)(255 * (1 - (double)(transparency / 100m)));
            paint.Color = paint.Color.WithAlpha(alpha);
        }

        canvas.DrawBitmap(bitmap, srcRect, destRect, paint);
        canvas.Restore();
    }

    private static void ApplyRotation(
        SKCanvas canvas,
        IShape shape,
        decimal x,
        decimal y,
        decimal width,
        decimal height)
    {
        const double epsilon = 1e-6;
        if (Math.Abs(shape.Rotation) > epsilon)
        {
            var centerX = x + (width / 2);
            var centerY = y + (height / 2);
            canvas.RotateDegrees(
                (float)shape.Rotation,
                (float)new Points(centerX).AsPixels(),
                (float)new Points(centerY).AsPixels()
            );
        }
    }
}
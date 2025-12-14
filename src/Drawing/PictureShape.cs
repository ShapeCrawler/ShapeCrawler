using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class PictureShape(Picture picture, P.Picture pPicture) : Shape(new Position(pPicture),
    new ShapeSize(pPicture), new ShapeId(pPicture), pPicture)
{
    public override decimal X
    {
        get => this.AbsoluteX();
        set
        {
            base.X = value;
            this.UpdateParentGroupX();
        }
    }

    public override decimal Y
    {
        get => this.AbsoluteY();
        set
        {
            base.Y = value;
            this.UpdateParentGroupY();
        }
    }

    public override decimal Width
    {
        get => this.AbsoluteWidth();
        set
        {
            base.Width = value;
            this.UpdateParentGroupWidth();
        }
    }

    public override decimal Height
    {
        get => this.AbsoluteHeight();
        set => base.Height = value;
    }

    public override IPicture Picture => picture;

    public override void CopyTo(P.ShapeTree pShapeTree) => picture.CopyTo(pShapeTree);

    internal override void Render(SKCanvas canvas)
    {
        if (picture.Image == null)
        {
            return;
        }
        
        var imageBytes = picture.Image.AsByteArray();
        using var bitmap = SKBitmap.Decode(imageBytes);
        var x = new Points(this.X).AsPixels();
        var y = new Points(this.Y).AsPixels();
        var width = new Points(this.Width).AsPixels();
        var height = new Points(this.Height).AsPixels();

        canvas.Save();
        ApplyRotation(canvas);

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

    private void ApplyRotation(SKCanvas canvas)
    {
        const double epsilon = 1e-6;
        if (Math.Abs(this.Rotation) <= epsilon)
        {
            return;
        }

        var centerX = this.X + (this.Width / 2);
        var centerY = this.Y + (this.Height / 2);
        canvas.RotateDegrees(
            (float)this.Rotation,
            (float)new Points(centerX).AsPixels(),
            (float)new Points(centerY).AsPixels()
        );
    }

    private decimal AbsoluteX()
    {
        var pGroupShapes = pPicture.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return base.X;
        }

        decimal absoluteX = base.X;

        foreach (var pGroupShape in pGroupShapes)
        {
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childOffset = transformGroup.ChildOffset!;
            var childExtents = transformGroup.ChildExtents!;
            var offset = transformGroup.Offset!;
            var extents = transformGroup.Extents!;

            decimal scaleFactor = 1.0m;
            if (childExtents.Cx!.Value != 0)
            {
                scaleFactor = (decimal)extents.Cx!.Value / childExtents.Cx!.Value;
            }

            var childOffsetX = new Emus(childOffset.X!.Value).AsPoints();
            absoluteX = ((absoluteX - childOffsetX) * scaleFactor) + new Emus(offset.X!.Value).AsPoints();
        }

        return absoluteX;
    }

    private decimal AbsoluteY()
    {
        var pGroupShapes = pPicture.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return base.Y;
        }

        decimal absoluteY = base.Y;

        foreach (var pGroupShape in pGroupShapes)
        {
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childOffset = transformGroup.ChildOffset!;
            var childExtents = transformGroup.ChildExtents!;
            var offset = transformGroup.Offset!;
            var extents = transformGroup.Extents!;

            decimal scaleFactor = 1.0m;
            if (childExtents.Cy!.Value != 0)
            {
                scaleFactor = (decimal)extents.Cy!.Value / childExtents.Cy!.Value;
            }

            var childOffsetY = new Emus(childOffset.Y!.Value).AsPoints();
            absoluteY = ((absoluteY - childOffsetY) * scaleFactor) + new Emus(offset.Y!.Value).AsPoints();
        }

        return absoluteY;
    }

    private decimal AbsoluteWidth()
    {
        var pGroupShapes = pPicture.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return base.Width;
        }

        decimal cumulativeScaleFactor = 1.0m;
        foreach (var pGroupShape in pGroupShapes)
        {
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childExtentsWidth = transformGroup.ChildExtents!.Cx!.Value;
            var extentsWidth = transformGroup.Extents!.Cx!.Value;
            if (childExtentsWidth == 0)
            {
                continue;
            }

            var scaleFactor = (decimal)extentsWidth / childExtentsWidth;
            cumulativeScaleFactor *= scaleFactor;
        }

        return base.Width * cumulativeScaleFactor;
    }

    private decimal AbsoluteHeight()
    {
        var pGroupShapes = pPicture.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return base.Height;
        }

        decimal cumulativeScaleFactor = 1.0m;
        foreach (var pGroupShape in pGroupShapes)
        {
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childExtentsHeight = transformGroup.ChildExtents!.Cy!.Value;
            var extentsHeight = transformGroup.Extents!.Cy!.Value;
            if (childExtentsHeight == 0)
            {
                continue;
            }

            var scaleFactor = (decimal)extentsHeight / childExtentsHeight;
            cumulativeScaleFactor *= scaleFactor;
        }

        return base.Height * cumulativeScaleFactor;
    }

    private void UpdateParentGroupX()
    {
        var pGroupShape = pPicture.Ancestors<P.GroupShape>().FirstOrDefault();
        if (pGroupShape is null)
        {
            return;
        }

        var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
        var aOffset = aTransformGroup.Offset!;
        var aExtents = aTransformGroup.Extents!;
        var aChildOffset = aTransformGroup.ChildOffset!;
        var aChildExtents = aTransformGroup.ChildExtents!;
        var groupedShapeXEmus = new Points(this.X).AsEmus();
        var groupShapeXEmus = aOffset.X!;

        if (groupedShapeXEmus < groupShapeXEmus)
        {
            var diff = groupShapeXEmus - groupedShapeXEmus;
            aOffset.X = new Int64Value(aOffset.X! - diff);
            aExtents.Cx = new Int64Value(aExtents.Cx! + diff);
            aChildOffset.X = new Int64Value(aChildOffset.X! - diff);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diff);

            return;
        }

        var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
        var groupedRightEmu = new Points(base.X + base.Width).AsEmus();
        if (groupedRightEmu > groupRightEmu)
        {
            var diffEmu = groupedRightEmu - groupRightEmu;
            aExtents.Cx = new Int64Value(aExtents.Cx! + diffEmu);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diffEmu);
        }
    }

    private void UpdateParentGroupY()
    {
        var pGroupShape = pPicture.Ancestors<P.GroupShape>().FirstOrDefault();
        if (pGroupShape is null)
        {
            return;
        }

        var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
        var aOffset = aTransformGroup.Offset!;
        var aExtents = aTransformGroup.Extents!;
        var aChildOffset = aTransformGroup.ChildOffset!;
        var aChildExtents = aTransformGroup.ChildExtents!;
        var groupedYEmus = new Points(this.Y).AsEmus();
        var groupYEmus = aOffset.Y!;
        if (groupedYEmus < groupYEmus)
        {
            var diff = groupYEmus - groupedYEmus;
            aOffset.Y = new Int64Value(aOffset.Y! - diff);
            aExtents.Cy = new Int64Value(aExtents.Cy! + diff);
            aChildOffset.Y = new Int64Value(aChildOffset.Y! - diff);
            aChildExtents.Cy = new Int64Value(aChildExtents.Cy! + diff);

            return;
        }

        var groupBottomEmu = aOffset.Y!.Value + aExtents.Cy!.Value;
        var groupedBottomEmu = groupedYEmus + new Points(base.Height).AsEmus();
        if (groupedBottomEmu > groupBottomEmu)
        {
            var diffEmu = groupedBottomEmu - groupBottomEmu;
            aExtents.Cy = new Int64Value(aExtents.Cy! + diffEmu);
            aChildExtents.Cy = new Int64Value(aChildExtents.Cy! + diffEmu);
        }
    }

    private void UpdateParentGroupWidth()
    {
        var pGroupShape = pPicture.Ancestors<P.GroupShape>().FirstOrDefault();
        if (pGroupShape is null)
        {
            return;
        }

        var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
        var aOffset = aTransformGroup.Offset!;
        var aExtents = aTransformGroup.Extents!;
        var aChildExtents = aTransformGroup.ChildExtents!;
        var groupedShapeWidthEmus = new Points(this.Width).AsEmus();
        var groupShapeWidthEmus = aExtents.Cx!;

        if (groupedShapeWidthEmus < groupShapeWidthEmus)
        {
            var diff = groupShapeWidthEmus - groupedShapeWidthEmus;
            aExtents.Cx = new Int64Value(aExtents.Cx! - diff);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx! - diff);

            return;
        }

        var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
        var groupedRightEmu = new Points(base.X + base.Width).AsEmus();
        if (groupedRightEmu > groupRightEmu)
        {
            var diffEmu = groupedRightEmu - groupRightEmu;
            aExtents.Cx = new Int64Value(aExtents.Cx! + diffEmu);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diffEmu);
        }
    }
}
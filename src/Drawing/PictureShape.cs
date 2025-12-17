using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class PictureShape(Picture picture, P.Picture pPicture) : DrawingShape(new Position(pPicture),
    new ShapeSize(pPicture), new ShapeId(pPicture), pPicture)
{
    public override decimal X
    {
        get => this.AbsoluteX();
        set
        {
            base.X = this.LocalX(value);
            this.UpdateParentGroupX();
        }
    }

    public override decimal Y
    {
        get => this.AbsoluteY();
        set
        {
            base.Y = this.LocalY(value);
            this.UpdateParentGroupY();
        }
    }

    public override decimal Width
    {
        get => this.AbsoluteWidth();
        set
        {
            base.Width = this.LocalWidth(value);
            this.UpdateParentGroupWidth();
        }
    }

    public override decimal Height
    {
        get => this.AbsoluteHeight();
        set => base.Height = this.LocalHeight(value);
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

    private static long ChildDiff(long parentDiff, long extents, long childExtents)
    {
        if (parentDiff == 0)
        {
            return 0;
        }

        if (childExtents == 0)
        {
            return parentDiff;
        }

        var scaleFactor = (decimal)extents / childExtents;
        if (scaleFactor == 0)
        {
            return parentDiff;
        }

        return (long)decimal.Round(parentDiff / scaleFactor, 0, MidpointRounding.AwayFromZero);
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

    private decimal LocalX(decimal absoluteX)
    {
        var pGroupShapes = pPicture.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return absoluteX;
        }

        var localX = absoluteX;
        for (var i = pGroupShapes.Length - 1; i >= 0; i--)
        {
            var pGroupShape = pGroupShapes[i];
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

            if (scaleFactor == 0)
            {
                scaleFactor = 1.0m;
            }

            var childOffsetX = new Emus(childOffset.X!.Value).AsPoints();
            var offsetX = new Emus(offset.X!.Value).AsPoints();
            localX = ((localX - offsetX) / scaleFactor) + childOffsetX;
        }

        return localX;
    }

    private decimal LocalY(decimal absoluteY)
    {
        var pGroupShapes = pPicture.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return absoluteY;
        }

        var localY = absoluteY;
        for (var i = pGroupShapes.Length - 1; i >= 0; i--)
        {
            var pGroupShape = pGroupShapes[i];
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

            if (scaleFactor == 0)
            {
                scaleFactor = 1.0m;
            }

            var childOffsetY = new Emus(childOffset.Y!.Value).AsPoints();
            var offsetY = new Emus(offset.Y!.Value).AsPoints();
            localY = ((localY - offsetY) / scaleFactor) + childOffsetY;
        }

        return localY;
    }

    private decimal LocalWidth(decimal absoluteWidth)
    {
        var scaleFactor = ShapePositionHelper.CalculateAbsoluteDimension(
            1.0m,
            pPicture,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.ChildExtents!.Cx!.Value,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.Extents!.Cx!.Value
        );

        if (scaleFactor == 0)
        {
            return absoluteWidth;
        }

        return absoluteWidth / scaleFactor;
    }

    private decimal LocalHeight(decimal absoluteHeight)
    {
        var scaleFactor = ShapePositionHelper.CalculateAbsoluteDimension(
            1.0m,
            pPicture,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.ChildExtents!.Cy!.Value,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.Extents!.Cy!.Value
        );

        if (scaleFactor == 0)
        {
            return absoluteHeight;
        }

        return absoluteHeight / scaleFactor;
    }

    private decimal AbsoluteWidth()
    {
        return ShapePositionHelper.CalculateAbsoluteDimension(
            base.Width,
            pPicture,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.ChildExtents!.Cx!.Value,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.Extents!.Cx!.Value
        );
    }

    private decimal AbsoluteHeight()
    {
        return ShapePositionHelper.CalculateAbsoluteDimension(
            base.Height,
            pPicture,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.ChildExtents!.Cy!.Value,
            groupShape => groupShape.GroupShapeProperties!.TransformGroup!.Extents!.Cy!.Value
        );
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
        var groupShapeXEmus = aOffset.X!.Value;

        if (groupedShapeXEmus < groupShapeXEmus)
        {
            var diffParent = groupShapeXEmus - groupedShapeXEmus;
            var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
            aOffset.X = new Int64Value(aOffset.X!.Value - diffParent);
            aExtents.Cx = new Int64Value(aExtents.Cx!.Value + diffParent);
            aChildOffset.X = new Int64Value(aChildOffset.X!.Value - diffChild);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value + diffChild);

            return;
        }

        var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
        var groupedRightEmu = new Points(this.X + this.Width).AsEmus();
        if (groupedRightEmu > groupRightEmu)
        {
            var diffParent = groupedRightEmu - groupRightEmu;
            var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
            aExtents.Cx = new Int64Value(aExtents.Cx!.Value + diffParent);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value + diffChild);
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
        var groupYEmus = aOffset.Y!.Value;
        if (groupedYEmus < groupYEmus)
        {
            var diffParent = groupYEmus - groupedYEmus;
            var diffChild = ChildDiff(diffParent, aExtents.Cy!.Value, aChildExtents.Cy!.Value);
            aOffset.Y = new Int64Value(aOffset.Y!.Value - diffParent);
            aExtents.Cy = new Int64Value(aExtents.Cy!.Value + diffParent);
            aChildOffset.Y = new Int64Value(aChildOffset.Y!.Value - diffChild);
            aChildExtents.Cy = new Int64Value(aChildExtents.Cy!.Value + diffChild);

            return;
        }

        var groupBottomEmu = aOffset.Y!.Value + aExtents.Cy!.Value;
        var groupedBottomEmu = new Points(this.Y + this.Height).AsEmus();
        if (groupedBottomEmu > groupBottomEmu)
        {
            var diffParent = groupedBottomEmu - groupBottomEmu;
            var diffChild = ChildDiff(diffParent, aExtents.Cy!.Value, aChildExtents.Cy!.Value);
            aExtents.Cy = new Int64Value(aExtents.Cy!.Value + diffParent);
            aChildExtents.Cy = new Int64Value(aChildExtents.Cy!.Value + diffChild);
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
        var groupShapeWidthEmus = aExtents.Cx!.Value;

        if (groupedShapeWidthEmus < groupShapeWidthEmus)
        {
            var diffParent = groupShapeWidthEmus - groupedShapeWidthEmus;
            var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
            aExtents.Cx = new Int64Value(aExtents.Cx!.Value - diffParent);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value - diffChild);

            return;
        }

        var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
        var groupedRightEmu = new Points(this.X + this.Width).AsEmus();
        if (groupedRightEmu > groupRightEmu)
        {
            var diffParent = groupedRightEmu - groupRightEmu;
            var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
            aExtents.Cx = new Int64Value(aExtents.Cx!.Value + diffParent);
            aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value + diffChild);
        }
    }

    // Helper for absolute position/size calculations
    private static class ShapePositionHelper
    {
        public static decimal CalculateAbsoluteDimension(
            decimal baseValue,
            OpenXmlElement shapeElement,
            Func<P.GroupShape, long> getChildExtents,
            Func<P.GroupShape, long> getExtents)
        {
            var pGroupShapes = shapeElement.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return baseValue;
            }

            decimal cumulativeScaleFactor = 1.0m;
            foreach (var pGroupShape in pGroupShapes)
            {
                var childExtents = getChildExtents(pGroupShape);
                var extents = getExtents(pGroupShape);
                if (childExtents == 0)
                {
                    continue;
                }

                var scaleFactor = (decimal)extents / childExtents;
                cumulativeScaleFactor *= scaleFactor;
            }

            return baseValue * cumulativeScaleFactor;
        }
    }
}
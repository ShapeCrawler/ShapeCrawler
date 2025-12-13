using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Groups;

internal sealed class GroupedShape : Shape
{
    private readonly P.Shape pShape;

    internal GroupedShape(Position position, ShapeSize shapeSize, ShapeId shapeId, P.Shape pShape) 
        : base(position, shapeSize, shapeId, pShape)
    {
        this.pShape = pShape;
    }

    public override decimal X
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = this.pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.X;
            }

            decimal absoluteX = base.X;

            // Apply the formula for each parent group in the hierarchy, from innermost to outermost
            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childOffset = transformGroup.ChildOffset!;
                var childExtents = transformGroup.ChildExtents!;
                var offset = transformGroup.Offset!;
                var extents = transformGroup.Extents!;

                // Calculate scale factor (ratio of group extents to child extents)
                decimal scaleFactor = 1.0m;
                if (childExtents.Cx!.Value != 0)
                {
                    scaleFactor = (decimal)extents.Cx!.Value / childExtents.Cx!.Value;
                }

                // Apply the formula: (childOffset - groupChildOffset) * scaleFactor + groupOffset
                var childOffsetX = new Emus(childOffset.X!.Value).AsPoints();
                absoluteX = ((absoluteX - childOffsetX) * scaleFactor) + new Emus(offset.X!.Value).AsPoints();
            }

            return absoluteX;
        }

        set
        {
            base.X = value;
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var groupedShapeXEmus = new Points(value).AsEmus();
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
    }

    public override decimal Y
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = this.pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.Y;
            }

            // Start with the shape's relative Y coordinate
            decimal absoluteY = base.Y;

            // Apply the formula for each parent group in the hierarchy, from innermost to outermost
            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childOffset = transformGroup.ChildOffset!;
                var childExtents = transformGroup.ChildExtents!;
                var offset = transformGroup.Offset!;
                var extents = transformGroup.Extents!;

                // Calculate scale factor (ratio of group extents to child extents)
                decimal scaleFactor = 1.0m;
                if (childExtents.Cy!.Value != 0)
                {
                    scaleFactor = (decimal)extents.Cy!.Value / childExtents.Cy!.Value;
                }

                // Apply the formula: (childOffset - groupChildOffset) * scaleFactor + groupOffset
                var childOffsetY = new Emus(childOffset.Y!.Value).AsPoints();
                absoluteY = ((absoluteY - childOffsetY) * scaleFactor) + new Emus(offset.Y!.Value).AsPoints();
            }

            return absoluteY;
        }

        set
        {
            base.Y = value;
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var groupedYEmu = new Points(value).AsEmus();
            var groupYEmu = aOffset.Y!;
            if (groupedYEmu < groupYEmu)
            {
                var diff = groupYEmu - groupedYEmu;
                aOffset.Y = new Int64Value(aOffset.Y! - diff);
                aExtents.Cy = new Int64Value(aExtents.Cy! + diff);
                aChildOffset.Y = new Int64Value(aChildOffset.Y! - diff);
                aChildExtents.Cy = new Int64Value(aChildExtents.Cy! + diff);

                return;
            }

            var groupBottomEmu = aOffset.Y!.Value + aExtents.Cy!.Value;
            var groupedBottomEmu = groupedYEmu + new Points(this.Height).AsEmus();
            if (groupedBottomEmu > groupBottomEmu)
            {
                var diffEmu = groupedBottomEmu - groupBottomEmu;
                aExtents.Cy = new Int64Value(aExtents.Cy! + diffEmu);
                aChildExtents.Cy = new Int64Value(aChildExtents.Cy! + diffEmu);
            }
        }
    }

    public override decimal Width
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = this.pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.Width;
            }

            // Calculate cumulative scale factor through all parent groups
            decimal cumulativeScaleFactor = 1.0m;

            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childExtentsWidth = transformGroup.ChildExtents!.Cx!.Value;
                var extentsWidth = transformGroup.Extents!.Cx!.Value;

                // Skip if either value is zero to avoid division by zero
                if (childExtentsWidth == 0)
                {
                    continue;
                }

                var scaleFactor = (decimal)extentsWidth / childExtentsWidth;
                cumulativeScaleFactor *= scaleFactor;
            }

            return base.Width * cumulativeScaleFactor;
        }

        set
        {
            base.Width = value;
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var groupedShapeWidthEmus = new Points(value).AsEmus();
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

    public override decimal Height
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = this.pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.Height;
            }

            // Calculate cumulative scale factor through all parent groups
            decimal cumulativeScaleFactor = 1.0m;

            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childExtentsCy = transformGroup.ChildExtents!.Cy!.Value;
                var extentsCy = transformGroup.Extents!.Cy!.Value;

                // Skip if either value is zero to avoid division by zero
                if (childExtentsCy == 0)
                {
                    continue;
                }

                var scaleFactor = (decimal)extentsCy / childExtentsCy;
                cumulativeScaleFactor *= scaleFactor;
            }

            return base.Height * cumulativeScaleFactor;
        }
        set => base.Height = value;
    }
}
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Groups;

internal sealed class GroupedShape(P.Shape pShape, Shape shape) : IShape
{
    public decimal X
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToList();
            if (pGroupShapes.Count == 0)
                return shape.X;

            // Start with the shape's relative X coordinate
            decimal absoluteX = shape.X;
            
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
                if (childExtents.Cx!.Value != 0) // Avoid division by zero
                {
                    scaleFactor = (decimal)extents.Cx!.Value / childExtents.Cx!.Value;
                }
                
                // Apply the formula: (childOffset - groupChildOffset) * scaleFactor + groupOffset
                var childOffsetX = new Emus(childOffset.X!.Value).AsPoints();
                absoluteX = (absoluteX - childOffsetX) * scaleFactor + new Emus(offset.X!.Value).AsPoints();
            }
            
            return absoluteX;
        }

        set
        {
            shape.X = value;
            var pGroupShape = pShape.Ancestors<P.GroupShape>().First();
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
            var groupedRightEmu = new Points(shape.X + shape.Width).AsEmus();
            if (groupedRightEmu > groupRightEmu)
            {
                var diffEmu = groupedRightEmu - groupRightEmu;
                aExtents.Cx = new Int64Value(aExtents.Cx! + diffEmu);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diffEmu);
            }
        }
    }

    public decimal Y
    {
        get => shape.Y;
        set
        {
            shape.Y = value;
            var pGroupShape = pShape.Ancestors<P.GroupShape>().First();
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

    public decimal Width
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToList();
            if (pGroupShapes.Count == 0)
                return shape.Width;

            // Calculate cumulative scale factor through all parent groups
            decimal cumulativeScaleFactor = 1.0m;
            
            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childExtentsWidth = transformGroup.ChildExtents!.Cx!.Value;
                var extentsWidth = transformGroup.Extents!.Cx!.Value;
                
                // Skip if either value is zero to avoid division by zero
                if (childExtentsWidth == 0)
                    continue;
                    
                var scaleFactor = (decimal)extentsWidth / childExtentsWidth;
                cumulativeScaleFactor *= scaleFactor;
            }
            
            return shape.Width * cumulativeScaleFactor;
        }
        set
        {
            shape.Width = value;
            var pGroupShape = pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
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
            var groupedRightEmu = new Points(shape.X + shape.Width).AsEmus();
            if (groupedRightEmu > groupRightEmu)
            {
                var diffEmu = groupedRightEmu - groupRightEmu;
                aExtents.Cx = new Int64Value(aExtents.Cx! + diffEmu);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diffEmu);
            }
        }
    }

    public decimal Height
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToList();
            if (pGroupShapes.Count == 0)
                return shape.Height;

            // Calculate cumulative scale factor through all parent groups
            decimal cumulativeScaleFactor = 1.0m;
            
            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childExtentsCy = transformGroup.ChildExtents!.Cy!.Value;
                var extentsCy = transformGroup.Extents!.Cy!.Value;
                
                // Skip if either value is zero to avoid division by zero
                if (childExtentsCy == 0)
                    continue;
                    
                var scaleFactor = (decimal)extentsCy / childExtentsCy;
                cumulativeScaleFactor *= scaleFactor;
            }
            
            return shape.Height * cumulativeScaleFactor;
        }
        set => shape.Height = value;
    }

    public int Id => shape.Id;

    public string Name
    {
        get => shape.Name;
        set => shape.Name = value;
    }

    public string AltText
    {
        get => shape.AltText;
        set => shape.AltText = value;
    }

    public bool Hidden => shape.Hidden;

    public PlaceholderType? PlaceholderType => shape.PlaceholderType;

    public Geometry GeometryType
    {
        get => shape.GeometryType;
        set => shape.GeometryType = value;
    }

    public decimal CornerSize
    {
        get => shape.CornerSize;
        set => shape.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => shape.Adjustments;
        set => shape.Adjustments = value;
    }

    public string? CustomData
    {
        get => shape.CustomData;
        set => shape.CustomData = value;
    }

    public ShapeContent ShapeContent => shape.ShapeContent;

    public bool HasOutline => shape.HasOutline;

    public IShapeOutline Outline => shape.Outline;

    public bool HasFill => shape.HasFill;

    public IShapeFill Fill => shape.Fill;

    public ITextBox? TextBox => shape.TextBox;

    public double Rotation => shape.Rotation;

    public bool Removable => shape.Removable;

    public string SDKXPath => shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;

    public IPresentation Presentation => shape.Presentation;

    public bool IsGroup => shape.IsGroup;

    public void Remove() => shape.Remove();

    public ITable AsTable() => shape.AsTable();

    public IMediaShape AsMedia() => shape.AsMedia();

    public void Duplicate() => throw new SCException("Duplicating grouped shape is not supported");
}
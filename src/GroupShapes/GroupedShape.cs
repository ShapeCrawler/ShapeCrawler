using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.GroupShapes;

internal sealed class GroupedShape : IShape
{
    private readonly P.Shape pShape;
    private readonly Shape shape;

    internal GroupedShape(P.Shape pShape, Shape shape)
    {
        this.pShape = pShape;
        this.shape = shape;
    }

    public decimal X
    {
        get
        {
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            
            var xGroupShapePt = new Emus(aOffset.X!).AsPoints();
            var groupShapeChildXPt = new Emus(pGroupShape.GroupShapeProperties!.TransformGroup!.ChildOffset!.X!.Value).AsPoints();
            var groupedShapeXPt = this.shape.X;
            
            return xGroupShapePt - (groupShapeChildXPt - groupedShapeXPt);
        }
        
        set
        {
            this.shape.X = value;
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
            var groupedRightEmu = new Points(this.shape.X + this.shape.Width).AsEmus();
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
        get => this.shape.Y;
        set
        {
            this.shape.Y = value;
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

    #region Decorated Shape

    public decimal Width
    {
        get => this.shape.Width;
        set => this.shape.Width = value;
    }

    public decimal Height
    {
        get => this.shape.Height;
        set => this.shape.Height = value;
    }

    public int Id => this.shape.Id;

    public string Name
    {
        get => this.shape.Name;
        set => this.shape.Name = value;
    }

    public string AltText
    {
        get => this.shape.AltText;
        set => this.shape.AltText = value;
    }

    public bool Hidden => this.shape.Hidden;
    
    public bool IsPlaceholder => this.shape.IsPlaceholder;
    
    public PlaceholderType PlaceholderType => this.shape.PlaceholderType;
   
    public Geometry GeometryType
    {
        get => this.shape.GeometryType;
        set => this.shape.GeometryType = value;
    }

    public decimal CornerSize
    {
        get => this.shape.CornerSize;
        set => this.shape.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => this.shape.Adjustments;
        set => this.shape.Adjustments = value;
    }

    public string? CustomData
    {
        get => this.shape.CustomData;
        set => this.shape.CustomData = value;
    }

    public ShapeContent ShapeContent => this.shape.ShapeContent;
    
    public bool HasOutline => this.shape.HasOutline;
    
    public IShapeOutline Outline => this.shape.Outline;
    
    public bool HasFill => this.shape.HasFill;
    
    public IShapeFill Fill => this.shape.Fill;
    
    public ITextBox? TextBox => this.shape.TextBox;
    
    public double Rotation => this.shape.Rotation;
    
    public bool Removeable => this.shape.Removeable;
    
    public string SDKXPath => this.shape.SDKXPath;
    
    public OpenXmlElement SDKOpenXmlElement => this.shape.SDKOpenXmlElement.CloneNode(true);

    public IPresentation Presentation => this.shape.Presentation;

    public void Remove() => this.shape.Remove();
    
    public ITable AsTable() => this.shape.AsTable();
    
    public IMediaShape AsMedia() => this.shape.AsMedia();

    public void Duplicate() => throw new SCException("Duplicating grouped shape is not supported");

    #endregion Decorated Shape
}
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection.GroupShapes;

internal sealed class GroupedShape : IShape
{
    private readonly P.Shape pShape;
    private readonly AutoShape decoratedShape;

    internal GroupedShape(OpenXmlPart sdkTypedOpenXmlPart, P.Shape pShape, AutoShape decoratedShape)
    {
        this.pShape = pShape;
        this.decoratedShape = decoratedShape;
    }

    public decimal X
    {
        get
        {
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            
            var xGroupShapeEmu = UnitConverter.HorizontalEmuToPixel(aOffset.X!);
            var groupShapeChildX = UnitConverter.HorizontalEmuToPixel(pGroupShape.GroupShapeProperties!.TransformGroup!.ChildOffset!.X!.Value);
            var groupedShapeX = this.decoratedShape.X;
            
            return xGroupShapeEmu - (groupShapeChildX - groupedShapeX);
        }
        
        set
        {
            this.decoratedShape.X = value;
            var xGroupedShapePx = value;
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var xGroupedShapeEmu = UnitConverter.HorizontalPixelToEmu(xGroupedShapePx);
            var xGroupShapeEmu = aOffset.X!;

            if (xGroupedShapeEmu < xGroupShapeEmu)
            {
                var diff = xGroupShapeEmu - xGroupedShapeEmu;
                aOffset.X = new Int64Value(aOffset.X! - diff);
                aExtents.Cx = new Int64Value(aExtents.Cx! + diff);
                aChildOffset.X = new Int64Value(aChildOffset.X! - diff);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diff);

                return;
            }

            var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
            var groupedRightEmu = UnitConverter.HorizontalPixelToEmu(this.decoratedShape.X + this.decoratedShape.Width);
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
        get => this.decoratedShape.Y;
        set
        {
            this.decoratedShape.Y = value;
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var groupedYEmu = UnitConverter.VerticalPixelToEmu(value);
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
            var groupedBottomEmu = groupedYEmu + UnitConverter.VerticalPixelToEmu(this.Height);
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
        get => this.decoratedShape.Width;
        set => this.decoratedShape.Width = value;
    }

    public decimal Height
    {
        get => this.decoratedShape.Height;
        set => this.decoratedShape.Height = value;
    }

    public int Id => this.decoratedShape.Id;

    public string Name => this.decoratedShape.Name;
    
    public bool Hidden => this.decoratedShape.Hidden;
    
    public bool IsPlaceholder => this.decoratedShape.IsPlaceholder;
    
    public PlaceholderType PlaceholderType => this.decoratedShape.PlaceholderType;
   
    public Geometry GeometryType => this.decoratedShape.GeometryType;

    public string? CustomData
    {
        get => this.decoratedShape.CustomData;
        set => this.decoratedShape.CustomData = value;
    }

    public ShapeType ShapeType => this.decoratedShape.ShapeType;
    
    public bool HasOutline => this.decoratedShape.HasOutline;
    
    public IShapeOutline Outline => this.decoratedShape.Outline;
    
    public bool HasFill => this.decoratedShape.HasFill;
    
    public IShapeFill Fill => this.decoratedShape.Fill;
    
    public bool IsTextHolder => this.decoratedShape.IsTextHolder;
    
    public ITextBox TextBox => this.decoratedShape.TextBox;
    
    public double Rotation => this.decoratedShape.Rotation;
    
    public bool Removeable => this.decoratedShape.Removeable;
    
    public string SdkXPath => this.decoratedShape.SdkXPath;
    
    public OpenXmlElement SdkOpenXmlElement => this.decoratedShape.SdkOpenXmlElement.CloneNode(true);

    public string Text
    {
        get => this.TextBox.Text;
        set => this.TextBox.Text = value;
    }

    public void Remove() => this.decoratedShape.Remove();
    
    public ITable AsTable() => this.decoratedShape.AsTable();
    
    public IMediaShape AsMedia() => this.decoratedShape.AsMedia();

    #endregion Decorated Shape
}
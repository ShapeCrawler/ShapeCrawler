using System;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.SlideShape;

internal sealed record GroupedSlideShape : IShape
{
    private readonly SlideShape slideShape;

    internal GroupedSlideShape(SlideShape slideShape)
    {
        this.slideShape = slideShape;
    }

    public int X
    {
        get => this.slideShape.X;
        set => this.UpdateX(value);
    }

    private void UpdateX(int value) => this.slideShape.X = value;

    public int Y 
    {
        get => this.slideShape.Y; 
        set => this.UpdateY(value);
    }

    private void UpdateY(int value)
    {
        this.slideShape.Y = value;
    }

    #region Properties
    public int Width
    {
        get => this.slideShape.Width; 
        set => this.slideShape.Width = value;
    }
    public int Height
    {
        get => this.slideShape.Height; 
        set => this.slideShape.Height = value;
    }
    public int Id => this.slideShape.Id;
    public string Name => this.slideShape.Name;
    public bool Hidden => this.slideShape.Hidden;

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder =>
        throw new SCException($"Grouped Shape cannot be a placeholder. Use {nameof(IShape.IsPlaceholder)} to check if the shape is a placeholder.");
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public bool HasOutline { get; }
    public IShapeOutline Outline => this.slideShape.Outline;
    public IShapeFill? Fill { get; }

    public bool IsTextHolder => false;

    public ITextFrame TextFrame =>
        throw new SCException($"The shape is not a text holder. Use {nameof(IShape.IsTextHolder)} property to check if the shape is a text holder.");

    public double Rotation { get; }
    #endregion Properties
}
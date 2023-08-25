using System;
using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Shapes;

internal sealed record GroupedSlideAutoShape : IAutoShape
{
    private readonly SlideAutoShape autoShape;
    private event EventHandler<int> XChanged;
    private event EventHandler<int> YChanged;

    internal GroupedSlideAutoShape(SlideAutoShape autoShape, EventHandler<int> xChangedHandler, EventHandler<int> yChangedHandler)
    {
        this.autoShape = autoShape;
        this.XChanged += xChangedHandler;
        this.YChanged += yChangedHandler;
    }

    public int X
    {
        get => this.autoShape.X;
        set => this.UpdateX(value);
    }

    private void UpdateX(int value)
    {
        this.autoShape.X = value;
        this.XChanged.Invoke(this, value);
    }

    public int Y 
    {
        get => this.autoShape.Y; 
        set => this.UpdateY(value);
    }

    private void UpdateY(int value)
    {
        this.autoShape.Y = value;
        this.YChanged.Invoke(this, value);
    }

    public int Width
    {
        get => this.autoShape.Width; 
        set => this.autoShape.Width = value;
    }

    public int Height
    {
        get => this.autoShape.Height; 
        set => this.autoShape.Height = value;
    }
    
    public int Id => this.autoShape.Id;
    public string Name => this.autoShape.Name;
    public bool Hidden => this.autoShape.Hidden;
    public bool IsPlaceholder() => this.autoShape.IsPlaceholder();

    public IPlaceholder Placeholder => this.autoShape.Placeholder;
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public IAutoShape? AsAutoShape()
    {
        throw new NotImplementedException();
    }

    public IShapeOutline Outline => this.autoShape.Outline;
    public IShapeFill? Fill { get; }
    public ITextFrame? TextFrame { get; }
    public bool IsTextHolder() => this.autoShape.IsTextHolder();

    public double Rotation { get; }

    public IAutoShape Duplicate()
    {
        throw new NotImplementedException();
    }
}
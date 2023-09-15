// ReSharper disable CheckNamespace

using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal record SlideOLEObject : IShape, IRemoveable
{
    private readonly SlidePart sdkSlidePart;
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly SimpleShape simpleShape;

    internal SlideOLEObject(SlidePart sdkSlidePart, P.GraphicFrame pGraphicFrame)
        : this(sdkSlidePart, pGraphicFrame, new SimpleShape(pGraphicFrame))
    {
    }

    private SlideOLEObject(SlidePart sdkSlidePart, P.GraphicFrame pGraphicFrame, SimpleShape simpleShape)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pGraphicFrame = pGraphicFrame;
        this.simpleShape = simpleShape;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
        this.Fill = new SlideShapeFill(sdkSlidePart, pGraphicFrame.Descendants<P.ShapeProperties>().First(), false);
    }

    public SCShapeType ShapeType => SCShapeType.OLEObject;
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public bool HasFill => true;
    public IShapeFill Fill { get; }
    
    #region SimpleShape
    
    public int X
    {
        get => this.simpleShape.X;
        set => this.simpleShape.X = value;
    }

    public int Y
    {
        get => this.simpleShape.Y;
        set => this.simpleShape.Y = value;
    }

    public int Width
    {
        get => this.simpleShape.Width;
        set => this.simpleShape.Width = value;
    }

    public int Height
    {
        get => this.simpleShape.Height;
        set => this.simpleShape.Height = value;
    }

    public int Id => this.simpleShape.Id;

    public string Name => this.simpleShape.Name;

    public bool Hidden => this.simpleShape.Hidden;
    public bool IsPlaceholder => this.simpleShape.IsPlaceholder;

    public IPlaceholder Placeholder => this.simpleShape.Placeholder;

    public SCGeometry GeometryType => this.simpleShape.GeometryType;
    
    public bool IsTextHolder => this.simpleShape.IsTextHolder;

    public string? CustomData
    {
        get => this.simpleShape.ParseCustomData();
        set => this.simpleShape.UpdateCustomData(value);
    }
    
    
    public ITextFrame TextFrame => this.simpleShape.TextFrame;
    public double Rotation => this.simpleShape.Rotation;
    public ITable AsTable() => this.simpleShape.AsTable();
    public IMediaShape AsMedia() => this.simpleShape.AsMedia();
    
    #endregion SimpleShape

    void IRemoveable.Remove()
    {
        this.pGraphicFrame.Remove();
    }
}
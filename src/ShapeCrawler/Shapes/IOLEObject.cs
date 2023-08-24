﻿// ReSharper disable CheckNamespace

using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal record SlideOLEObject : IShape
{
    private readonly Shape shape;

    internal SlideOLEObject(P.GraphicFrame pGraphicFrame, SlideShapes shapes, Shape shape)
    {
        this.shape = shape;
    }

    public int X
    {
        get => this.shape.X();
        set => this.shape.UpdateX(value);
    }

    public int Y
    {
        get => this.shape.Y();
        set => this.shape.UpdateY(value);
    }

    public int Width
    {
        get => this.shape.Width();
        set => this.shape.UpdateWidth(value);
    }

    public int Height
    {
        get => this.shape.Height();
        set => this.shape.UpdateHeight(value);
    }

    public int Id => this.shape.Id();

    public string Name => this.shape.Name();

    public bool Hidden => this.shape.Hidden();
    public bool IsPlaceholder() => false;

    public IPlaceholder Placeholder => new NullPlaceholder();

    public SCGeometry GeometryType => this.shape.GeometryType();

    public string? CustomData
    {
        get => this.shape.CustomData();
        set => this.shape.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.OLEObject;

    public IAutoShape? AsAutoShape()
    {
        throw new System.NotImplementedException();
    }

    internal void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new System.NotImplementedException();
    }

    internal string ToJson()
    {
        throw new System.NotImplementedException();
    }
}
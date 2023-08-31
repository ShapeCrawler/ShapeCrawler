using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed record LayoutShapeFill : IShapeFill
{
    private readonly SlideLayoutPart sdkLayoutPart;
    private readonly P.ShapeProperties pShapeProperties;
    private readonly BooleanValue useBgFill;

    internal LayoutShapeFill(SlideLayoutPart sdkLayoutPart, P.ShapeProperties pShapeProperties, BooleanValue useBgFill)
    {
        this.sdkLayoutPart = sdkLayoutPart;
        this.pShapeProperties = pShapeProperties;
        this.useBgFill = useBgFill;
    }

    public SCFillType Type { get; }
    public IImage? Picture { get; }
    public string? Color { get; }
    public double AlphaPercentage { get; }
    public double LuminanceModulation { get; }
    public double LuminanceOffset { get; }
    public void SetPicture(Stream image)
    {
        throw new NotImplementedException();
    }

    public void SetColor(string hex)
    {
        throw new NotImplementedException();
    }
}
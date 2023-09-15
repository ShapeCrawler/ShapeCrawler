using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed class SimpleShape : IShape
{
    private readonly OpenXmlElement pShapeTreeElement;
    private readonly Position position;
    private readonly ShapeSize size;
    private const string customDataElementName = "ctd";

    internal SimpleShape(OpenXmlElement pShapeTreeElement)
    {
        this.pShapeTreeElement = pShapeTreeElement;
        this.position = new Position(this.pShapeTreeElement);
        this.size = new ShapeSize(pShapeTreeElement);
    }

    internal string? ParseCustomData()
    {
        const string pattern = @$"<{customDataElementName}>(.*)<\/{customDataElementName}>";

#if NETSTANDARD2_0
        var regex = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(100));
#else
        var regex = new Regex(pattern, RegexOptions.NonBacktracking);
#endif

        var elementText = regex.Match(this.pShapeTreeElement.InnerXml).Groups[1];
        if (elementText.Value.Length == 0)
        {
            return null;
        }

        return elementText.Value;
    }

    internal void UpdateCustomData(string? value)
    {
        var customDataElement =
            $@"<{customDataElementName}>{value}</{customDataElementName}>";
        this.pShapeTreeElement.InnerXml += customDataElement;
    }

    public int X
    {
        get => this.position.X();
        set => this.position.UpdateX(value);
    }

    public int Y
    {
        get => this.position.Y();
        set => this.position.UpdateY(value);
    }

    public int Width
    {
        get => this.size.Width();
        set => this.size.UpdateWidth(value);
    }

    public int Height
    {
        get => this.size.Width();
        set => this.size.UpdateWidth(value);
    }

    public int Id => (int)this.pShapeTreeElement.GetNonVisualDrawingProperties().Id!.Value!;

    public string Name => this.pShapeTreeElement.GetNonVisualDrawingProperties().Name!.Value!;

    public bool Hidden
    {
        get
        {
            var parsedHiddenValue = this.pShapeTreeElement.GetNonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }
    }

    public bool IsPlaceholder => false;
    public IPlaceholder Placeholder => throw new SCException($"The shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if shape is a placeholder.");

    public SCGeometry GeometryType => SCGeometry.Rectangle;

    public string? CustomData
    {
        get => this.ParseCustomData();
        set => this.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.AutoShape;
    public bool HasOutline => false;
    public IShapeOutline Outline => throw new SCException($"Shape does not have outline. Use {nameof(IShape.HasOutline)} property to check if the shape has outline.");
    public bool HasFill => false;
    public IShapeFill Fill => throw new SCException($"Shape does not have fill. Use {nameof(IShape.HasFill)} property to check if the shape has fill.");
    public bool IsTextHolder => false;
    public ITextFrame TextFrame => throw new SCException($"The shape is not a text holder. Use {nameof(IShape.IsTextHolder)} method to check it.");
    public double Rotation => throw new NotImplementedException();
    public ITable AsTable() => throw new SCException($"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media (audio, video, etc.");
}
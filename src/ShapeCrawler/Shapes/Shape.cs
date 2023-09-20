using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal abstract class Shape : IShape
{
    private readonly Position position;
    private readonly ShapeSize size;
    private readonly ShapeId shapeId;
    private const string customDataElementName = "ctd";
    
    protected readonly OpenXmlElement pShapeTreeElement;

    internal Shape(OpenXmlElement pShapeTreeElement)
    {
        this.pShapeTreeElement = pShapeTreeElement;
        this.position = new Position(pShapeTreeElement);
        this.size = new ShapeSize(pShapeTreeElement);
        this.shapeId = new ShapeId(pShapeTreeElement);
    }

    private string? ParseCustomData()
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
        get => this.size.Height();
        set => this.size.UpdateWidth(value);
    }

    public int Id => this.shapeId.Value();

    public string Name => this.pShapeTreeElement.GetNonVisualDrawingProperties().Name!.Value!;

    public bool Hidden
    {
        get
        {
            var parsedHiddenValue = this.pShapeTreeElement.GetNonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }
    }

    public virtual bool IsPlaceholder => false;

    public virtual IPlaceholder Placeholder => throw new SCException(
        $"The shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if shape is a placeholder.");

    public virtual SCGeometry GeometryType => SCGeometry.Rectangle;

    public string? CustomData
    {
        get => this.ParseCustomData();
        set => this.UpdateCustomData(value);
    }

    public abstract SCShapeType ShapeType { get; }
    public virtual bool HasOutline => false;

    public virtual IShapeOutline Outline => throw new SCException(
        $"Shape does not have outline. Use {nameof(IShape.HasOutline)} property to check if the shape has outline.");

    public virtual bool HasFill => false;

    public virtual IShapeFill Fill =>
        throw new SCException(
            $"Shape does not have fill. Use {nameof(IShape.HasFill)} property to check if the shape has fill.");

    public virtual bool IsTextHolder => false;

    public virtual ITextFrame TextFrame =>
        throw new SCException($"The shape is not a text holder. Use {nameof(IShape.IsTextHolder)} method to check it.");

    public double Rotation => throw new NotImplementedException();

    public virtual ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public virtual IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media (audio, video, etc.");
}
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a slide number.
/// </summary>
public interface IMasterSlideNumber : IPosition
{
    /// <summary>
    ///     Gets font.
    /// </summary>
    ISlideNumberFont Font { get; }
}

internal sealed class MasterSlideNumber : IMasterSlideNumber
{
    private readonly Position position;

    internal MasterSlideNumber(P.Shape sdkPShape)
    :this(sdkPShape, new Position(sdkPShape))
    {
    }
    
    private MasterSlideNumber(P.Shape sdkPShape, Position position)
    {
        this.position = position;
        var aDefaultRunProperties =
            sdkPShape.TextBody!.ListStyle!.Level1ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() !;
        this.Font = new SlideNumberFont(aDefaultRunProperties);
    }

    public ISlideNumberFont Font { get; }

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
}
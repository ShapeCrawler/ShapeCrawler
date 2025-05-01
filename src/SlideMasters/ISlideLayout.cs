using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a Slide Layout.
/// </summary>
public interface ISlideLayout
{
    /// <summary>
    ///     Gets layout name.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets layout shape collection.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets slide master.
    /// </summary>
    ISlideMaster SlideMaster { get; }
}

internal sealed class SlideLayout : ISlideLayout
{
    private readonly SlideLayoutPart slideLayoutPart;

    internal SlideLayout(SlideLayoutPart sdkLayoutPart)
        : this(sdkLayoutPart, new SlideMaster(sdkLayoutPart.SlideMasterPart!))
    {
    }

    private SlideLayout(SlideLayoutPart sdkLayoutPart, ISlideMaster slideMaster)
    {
        this.slideLayoutPart = sdkLayoutPart;
        this.SlideMaster = slideMaster;
        this.Shapes = new ShapeCollection(this.slideLayoutPart);
    }

    public string Name => this.slideLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;
    
    public IShapeCollection Shapes { get; }
    
    public ISlideMaster SlideMaster { get; }
    
    internal SlideLayoutPart SdkSlideLayoutPart() => this.slideLayoutPart;
}
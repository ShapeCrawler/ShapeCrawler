using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;

namespace ShapeCrawler.SlideMasters;

/// <summary>
///     Represents a Slide Layout.
/// </summary>
public interface ISlideLayout
{
    /// <summary>
    ///     Gets parent Slide Master.
    /// </summary>
    ISlideMaster SlideMaster { get; }

    /// <summary>
    ///     Gets collection of shape.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets layout name.
    /// </summary>
    string Name { get; }
}

internal sealed class SCSlideLayout : SlideStructure, ISlideLayout
{
    private readonly ResettableLazy<ShapeCollection> shapes;
    private readonly SCSlideMaster slideMaster;

    internal SCSlideLayout(SCSlideMaster slideMaster, SlideLayoutPart slideLayoutPart, int number)
        : base(slideMaster.Presentation)
    {
        this.slideMaster = slideMaster;
        this.SlideLayoutPart = slideLayoutPart;
        this.shapes = new ResettableLazy<ShapeCollection>(() =>
            new ShapeCollection(slideLayoutPart, this));
        this.Number = number;
    }

    public string Name => this.GetName();

    public ISlideMaster SlideMaster => this.slideMaster;
    
    public override int Number { get; set; }

    public override IShapeCollection Shapes => this.shapes.Value;

    internal SlideLayoutPart SlideLayoutPart { get; }

    internal SCSlideMaster SlideMasterInternal => (SCSlideMaster)this.SlideMaster;

    internal ShapeCollection ShapesInternal => (ShapeCollection)this.Shapes;

    internal override TypedOpenXmlPart TypedOpenXmlPart => this.SlideLayoutPart;
    
    private string GetName()
    {
        return this.SlideLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;
    }
}
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;

namespace ShapeCrawler.SlideMasters;

/// <summary>
///     Represents a Slide Layout.
/// </summary>
[SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — Shape Crawler")]
internal class SCSlideLayout : SlideObject, ISlideLayout
{
    private readonly ResettableLazy<ShapeCollection> shapes;
    private readonly SCSlideMaster slideMaster;

    internal SCSlideLayout(SCSlideMaster slideMaster, SlideLayoutPart slideLayoutPart, int number)
    : base(slideMaster.Presentation)
    {
        this.slideMaster = slideMaster;
        this.SlideLayoutPart = slideLayoutPart;
        this.shapes = new ResettableLazy<ShapeCollection>(() =>
            ShapeCollection.Create(slideLayoutPart, this));
        this.Number = number;
    }

    public IShapeCollection Shapes => this.shapes.Value;

    public string Name => this.GetName();

    public ISlideMaster SlideMaster => this.slideMaster;
    
    public override int Number { get; set; }

    internal SlideLayoutPart SlideLayoutPart { get; }

    internal SCSlideMaster SlideMasterInternal => (SCSlideMaster)this.SlideMaster;

    internal ShapeCollection ShapesInternal => (ShapeCollection)this.Shapes;

    internal override TypedOpenXmlPart TypedOpenXmlPart => this.SlideLayoutPart;
    
    private string GetName()
    {
        return this.SlideLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;
    }
}
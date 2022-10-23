using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
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

    internal SCSlideLayout(SCSlideMaster slideMaster, SlideLayoutPart slideLayoutPart)
    : base(slideMaster.Presentation)
    {
        this.slideMaster = slideMaster;
        this.SlideLayoutPart = slideLayoutPart;
        this.shapes = new ResettableLazy<ShapeCollection>(() =>
            ShapeCollection.Create(slideLayoutPart, this));
    }

    public IShapeCollection Shapes => this.shapes.Value;

    public ISlideMaster SlideMaster => this.slideMaster;

    public SCPresentation PresentationInternal => this.SlideMasterInternal.PresentationInternal;

    internal SlideLayoutPart SlideLayoutPart { get; }

    internal SCSlideMaster SlideMasterInternal => (SCSlideMaster)this.SlideMaster;

    internal ShapeCollection ShapesInternal => (ShapeCollection)this.Shapes;

    internal override TypedOpenXmlPart TypedOpenXmlPart => this.SlideLayoutPart;
}
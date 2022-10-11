using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using OneOf;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Drawing;

/// <inheritdoc cref="IPicture" />
[SuppressMessage("ReSharper", "SuggestBaseTypeForParameterInConstructor", Justification = "Internal member")]
internal class SlidePicture : SlideShape, IPicture
{
    private readonly StringValue picReference;

    internal SlidePicture(P.Picture pPicture, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, StringValue picReference)
        : base(pPicture, slideObject, null)
    {
        this.picReference = picReference;
    }

    public IImage Image => SCImage.ForPicture(this, this.Slide.TypedOpenXmlPart, this.picReference);

    public override SCShapeType ShapeType => SCShapeType.Picture;
}
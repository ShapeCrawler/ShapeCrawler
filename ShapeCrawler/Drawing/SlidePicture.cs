using System;
using System.Diagnostics.CodeAnalysis;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Drawing;

/// <inheritdoc cref="IPicture" />
[SuppressMessage("ReSharper", "SuggestBaseTypeForParameterInConstructor", Justification = "Internal member")]
internal class SlidePicture : SlideShape, IPicture
{
    private readonly string blipEmbed;
    private readonly A.Blip aBlip;

    internal SlidePicture(P.Picture pPicture, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, A.Blip aBlip)
        : base(pPicture, slideObject, null)
    {
        this.aBlip = aBlip;
        this.blipEmbed = aBlip.Embed!.Value!;
    }

    public IImage Image => SCImage.ForPicture(this, this.Slide.TypedOpenXmlPart, this.blipEmbed);

    public string? SvgContent => this.GetSvgContent();

    public override SCShapeType ShapeType => SCShapeType.Picture;

    private string? GetSvgContent()
    {
        throw new NotImplementedException();
    }
}
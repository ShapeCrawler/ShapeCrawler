using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Pictures;

/// <inheritdoc cref="IPicture" />
[SuppressMessage("ReSharper", "SuggestBaseTypeForParameterInConstructor", Justification = "Internal member")]
internal sealed class SCPicture : SlideSCShape, IPicture
{
    private readonly StringValue? blipEmbed;
    private readonly A.Blip aBlip;

    internal SCPicture(P.Picture pPicture, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, A.Blip aBlip)
        : base(pPicture, slideObject, null)
    {
        this.aBlip = aBlip;
        this.blipEmbed = aBlip.Embed;
    }

    public IImage Image => SCImage.ForPicture(this, this.Slide.TypedOpenXmlPart, this.blipEmbed);

    public string? SvgContent => this.GetSvgContent();

    public override SCShapeType ShapeType => SCShapeType.Picture;

    internal override void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    private string? GetSvgContent()
    {
        var bel = this.aBlip.GetFirstChild<A.BlipExtensionList>();
        var svgBlipList = bel?.Descendants<SVGBlip>();
        if (svgBlipList == null)
        {
            return null;
        }

        var svgId = svgBlipList.First().Embed!.Value!;

        var imagePart = (ImagePart)this.Slide.TypedOpenXmlPart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }
}
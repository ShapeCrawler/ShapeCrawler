// ReSharper disable CheckNamespace

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Shapes;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

/// <summary>
///     Represents a picture shape on a slide.
/// </summary>
public interface IPicture : IShape
{
    /// <summary>
    ///     Gets image. Returns <see langword="null"/> if the picture is not binary picture. 
    /// </summary>
    IImage? Image { get; }

    /// <summary>
    ///     Gets SVG content. Returns <see langword="null"/> if the picture is not SVG graphic.
    /// </summary>
    string? SvgContent { get; }
}

internal sealed class SCPicture : SCShape, IPicture
{
    private readonly StringValue? blipEmbed;
    private readonly A.Blip aBlip;

    internal SCPicture(
        P.Picture pPicture, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection,
        A.Blip aBlip)
        : base(pPicture, parentSlideObject, parentShapeCollection)
    {
        this.aBlip = aBlip;
        this.blipEmbed = aBlip.Embed;
    }

    public IImage Image => SCImage.ForPicture(this, ((SlideStructure)this.SlideStructure).TypedOpenXmlPart, this.blipEmbed);

    public string? SvgContent => this.GetSvgContent();

    public override SCShapeType ShapeType => SCShapeType.Picture;

    internal override void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal override string ToJson()
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

        var slideStructureCore = (SlideStructure)this.SlideStructure;
        var imagePart = (ImagePart)slideStructureCore.TypedOpenXmlPart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }
}
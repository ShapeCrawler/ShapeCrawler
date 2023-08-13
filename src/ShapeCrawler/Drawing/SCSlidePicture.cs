using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class SCSlidePicture : IPicture
{
    private readonly StringValue? blipEmbed;
    private readonly A.Blip aBlip;

    internal SCSlidePicture(P.Picture pPicture, SCSlideShapes shapeCollection, A.Blip aBlip)
    {
        this.aBlip = aBlip;
        this.blipEmbed = aBlip.Embed;
    }

    public IImage Image =>
        SCImage.ForPicture(this.slideTypedOpenXmlPart, this.blipEmbed, this.imageParts);

    public string? SvgContent => this.GetSvgContent();

    public SCShapeType ShapeType => SCShapeType.Picture;

    /// <summary>
    ///     Copies all required parts from the source slide if they do not exist.
    /// </summary>
    /// <param name="sourceSlide">Source slide.</param>
    internal void CopyParts(ISlideStructure sourceSlide)
    {
        if (this.blipEmbed is null)
        {
            return;
        }

        if (this.slideTypedOpenXmlPart.GetPartById(this.blipEmbed.Value!) is not ImagePart imagePart)
        {
            return;
        }

        // Creates a new part in this slide with a new Id...
        var imgPartRId = this.slideTypedOpenXmlPart.GetNextRelationshipId();

        // Adds to current slide parts and update relation id.
        var nImagePart = this.slideTypedOpenXmlPart.AddNewPart<ImagePart>(imagePart.ContentType, imgPartRId);
        using var stream = imagePart.GetStream(FileMode.Open);
        stream.Position = 0;
        nImagePart.FeedData(stream);

        this.blipEmbed.Value = imgPartRId;
    }

    internal override void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal override IHtmlElement ToHtmlElement()
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

        var imagePart = (ImagePart)this.slideTypedOpenXmlPart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }
}
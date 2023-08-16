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

internal sealed record SCSlidePicture : SCSlideShape, IPicture
{
    private readonly StringValue blipEmbed;
    private readonly P.Picture pPicture;
    private readonly SCSlideShapes parentShapeCollection;
    private readonly A.Blip aBlip;
    private readonly Shape shape;

    internal SCSlidePicture(
        P.Picture pPicture,
        SCSlideShapes parentShapeCollection,
        A.Blip aBlip,
        Shape shape) : base(pPicture)
    {
        this.pPicture = pPicture;
        this.parentShapeCollection = parentShapeCollection;
        this.aBlip = aBlip;
        this.shape = shape;
        this.blipEmbed = aBlip.Embed!;
    }

    public IImage Image => new SCImage(this, this.blipEmbed.Value!);

    public string? SvgContent => this.GetSvgContent();

    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.shape.Id();
    public string Name => this.shape.Name();
    public bool Hidden { get; }
    public IPlaceholder? Placeholder { get; }
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => SCShapeType.Picture;
    public IAutoShape AsAutoShape()
    {
        throw new NotImplementedException();
    }

    internal void CopyParts()
    {
        if (this.blipEmbed is null)
        {
            return;
        }

        var sdkSlidePart = this.parentShapeCollection.SDKSlidePart();
        if (sdkSlidePart.GetPartById(this.blipEmbed.Value!) is not ImagePart imagePart)
        {
            return;
        }

        // Creates a new part in this slide with a new Id...
        var imgPartRId = sdkSlidePart.GetNextRelationshipId();

        // Adds to current slide parts and update relation id.
        var nImagePart = sdkSlidePart.AddNewPart<ImagePart>(imagePart.ContentType, imgPartRId);
        using var stream = imagePart.GetStream(FileMode.Open);
        stream.Position = 0;
        nImagePart.FeedData(stream);

        this.blipEmbed.Value = imgPartRId;
    }

    internal void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new NotImplementedException();
    }

    internal string ToJson()
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

        var imagePart = (ImagePart)this.SDKSlidePart().GetPartById(svgId);
        using var svgStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }

    internal SlidePart SDKSlidePart()
    {
        return this.parentShapeCollection.SDKSlidePart();
    }

    public int X { get; set; }
    public int Y { get; set; }
    
    internal override bool Copyable()
    {
        return true;
    }
}
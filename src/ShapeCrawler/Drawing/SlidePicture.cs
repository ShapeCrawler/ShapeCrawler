using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed record SlidePicture : IPicture, IRemoveable 
{
    private readonly StringValue blipEmbed;
    private readonly P.Picture pPicture;
    private readonly A.Blip aBlip;
    private readonly Shape shape;
    private readonly SlidePart sdkSlidePart;

    internal SlidePicture(SlidePart sdkSlidePart, P.Picture pPicture, A.Blip aBlip)
        : this(sdkSlidePart, pPicture, aBlip, new Shape(pPicture), new SlidePictureImage(sdkSlidePart, aBlip))
    {
    }

    private SlidePicture(SlidePart sdkSlidePart, P.Picture pPicture, A.Blip aBlip, Shape shape, IImage image)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pPicture = pPicture;
        this.aBlip = aBlip;
        this.shape = shape;
        this.Image = image;
        this.blipEmbed = aBlip.Embed!;
    }

    public IImage Image { get; }

    public string? SvgContent => this.GetSvgContent();

    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.shape.Id();
    public string Name => this.shape.Name();
    public bool Hidden { get; }

    public bool IsPlaceholder()
    {
        throw new NotImplementedException();
    }

    public IPlaceholder? Placeholder { get; }
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => SCShapeType.Picture;

    public IAutoShape AsAutoShape()
    {
        throw new NotImplementedException();
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

        var imagePart = (ImagePart)this.sdkSlidePart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }

    public int X { get; set; }
    public int Y { get; set; }

    internal void CopyTo(int id, P.ShapeTree pShapeTree, IEnumerable<string> existingShapeNames,
        SlidePart targetSdkSlidePart)
    {
        var copy = this.pPicture.CloneNode(true);
        copy.GetNonVisualDrawingProperties().Id = new UInt32Value((uint)id);
        pShapeTree.AppendChild(copy);
        var copyName = copy.GetNonVisualDrawingProperties().Name!.Value!;
        if (existingShapeNames.Any(existingShapeName => existingShapeName == copyName))
        {
            var currentShapeCollectionSuffixes = existingShapeNames
                .Where(c => c.StartsWith(copyName, StringComparison.InvariantCulture))
                .Select(c => c.Substring(copyName.Length))
                .ToArray();

            // We will try to check numeric suffixes only.
            var numericSuffixes = new List<int>();

            foreach (var currentSuffix in currentShapeCollectionSuffixes)
            {
                if (int.TryParse(currentSuffix, out var numericSuffix))
                {
                    numericSuffixes.Add(numericSuffix);
                }
            }

            numericSuffixes.Sort();
            var lastSuffix = numericSuffixes.LastOrDefault() + 1;
            copy.GetNonVisualDrawingProperties().Name = copyName + " " + lastSuffix;
        }

        // COPY PARTS
        var sourceSdkSlidePart = this.sdkSlidePart;
        var sourceImagePart = (ImagePart)sourceSdkSlidePart.GetPartById(this.blipEmbed.Value!);

        // Creates a new part in this slide with a new Id...
        var targetImagePartRId = targetSdkSlidePart.GetNextRelationshipId();

        // Adds to current slide parts and update relation id.
        var targetImagePart = targetSdkSlidePart.AddNewPart<ImagePart>(sourceImagePart.ContentType, targetImagePartRId);
        using var sourceImageStream = sourceImagePart.GetStream(FileMode.Open);
        sourceImageStream.Position = 0;
        targetImagePart.FeedData(sourceImageStream);

        copy.Descendants<A.Blip>().First().Embed = targetImagePartRId;
    }

    public void Remove()
    {
        this.pPicture.Remove();
    }
}
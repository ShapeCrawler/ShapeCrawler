using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed record SlidePicture : IPicture, IRemoveable 
{
    private readonly StringValue blipEmbed;
    private readonly P.Picture sdkPPicture;
    private readonly A.Blip sdkABlip;
    private readonly Shape shape;
    private readonly SlidePart sdkSlidePart;

    internal SlidePicture(SlidePart sdkSlidePart, P.Picture sdkPPicture, A.Blip sdkABlip)
        : this(sdkSlidePart, sdkPPicture, sdkABlip, new Shape(sdkPPicture), new SlidePictureImage(sdkSlidePart, sdkABlip))
    {
    }

    private SlidePicture(SlidePart sdkSlidePart, P.Picture sdkPPicture, A.Blip sdkABlip, Shape shape, IImage image)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.sdkPPicture = sdkPPicture;
        this.sdkABlip = sdkABlip;
        this.shape = shape;
        this.Image = image;
        this.blipEmbed = sdkABlip.Embed!;
        this.Outline = new SlideShapeOutline(sdkSlidePart, sdkPPicture.ShapeProperties!);
        this.Fill = new SlideShapeFill(sdkSlidePart, sdkPPicture.ShapeProperties!, false);
    }

    public IImage Image { get; }

    public string? SvgContent => this.GetSvgContent();

    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.shape.Id();
    public string Name => this.shape.Name();
    public bool Hidden { get; }
    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => throw new SCException(
        $"The Picture shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} to check if the shape is a placeholder.");
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => SCShapeType.Picture;
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public IShapeFill Fill { get; }
    public bool IsTextHolder { get; }

    public ITextFrame TextFrame => throw new SCException($"The Picture shape is not a text holder. Use {nameof(IShape.IsTextHolder)} method to check it.");
    public double Rotation { get; }

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
        var bel = this.sdkABlip.GetFirstChild<A.BlipExtensionList>();
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
        var copy = this.sdkPPicture.CloneNode(true);
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
        this.sdkPPicture.Remove();
    }
}
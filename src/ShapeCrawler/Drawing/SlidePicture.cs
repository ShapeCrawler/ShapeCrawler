using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class SlidePicture : CopyableShape, IPicture, IRemoveable 
{
    private readonly StringValue blipEmbed;
    private readonly P.Picture pPicture;
    private readonly A.Blip aBlip;
    private readonly SlidePart sdkSlidePart;

    internal SlidePicture(
        SlidePart sdkSlidePart, 
        P.Picture pPicture, 
        A.Blip aBlip)
        : this(sdkSlidePart, pPicture, aBlip, new SlidePictureImage(sdkSlidePart, aBlip))
    {
    }

    private SlidePicture(SlidePart sdkSlidePart, P.Picture pPicture, A.Blip aBlip, IImage image)
    :base(pPicture)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pPicture = pPicture;
        this.aBlip = aBlip;
        this.Image = image;
        this.blipEmbed = aBlip.Embed!;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pPicture.ShapeProperties!);
        this.Fill = new SlideShapeFill(sdkSlidePart, pPicture.ShapeProperties!, false);
    }

    public IImage Image { get; }
    public string? SvgContent => this.GetSvgContent();
    public override SCGeometry GeometryType => SCGeometry.Rectangle;
    public override SCShapeType ShapeType => SCShapeType.Picture;
    public override bool HasOutline => true;
    public override IShapeOutline Outline { get; }

    public override bool HasFill => true;
    public override IShapeFill Fill { get; }

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

    public void Remove()
    {
        this.pPicture.Remove();
    }

    internal override void CopyTo(int id, P.ShapeTree pShapeTree, IEnumerable<string> existingShapeNames, SlidePart targetSdkSlidePart)
    {
        base.CopyTo(id, pShapeTree, existingShapeNames, targetSdkSlidePart);

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

        var copy = this.pShapeTreeElement.CloneNode(true);
        copy.Descendants<A.Blip>().First().Embed = targetImagePartRId;
    }
}
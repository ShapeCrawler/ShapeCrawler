using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class Picture : CopyableShape, IPicture
{
    private readonly StringValue blipEmbed;
    private readonly P.Picture pPicture;
    private readonly A.Blip aBlip;
    private readonly ShapeGeometry shapeGeometry;

    internal Picture(
        OpenXmlPart openXmlPart,
        P.Picture pPicture,
        A.Blip aBlip)
        : this(openXmlPart, pPicture, aBlip, new SlidePictureImage(openXmlPart, aBlip))
    {
    }

    private Picture(OpenXmlPart openXmlPart, P.Picture pPicture, A.Blip aBlip, IImage image)
        : base(openXmlPart, pPicture)
    {
        this.pPicture = pPicture;
        this.aBlip = aBlip;
        this.Image = image;
        this.blipEmbed = aBlip.Embed!;
        this.Outline = new SlideShapeOutline(openXmlPart, pPicture.ShapeProperties!);
        this.Fill = new ShapeFill(openXmlPart, pPicture.ShapeProperties!);
        this.shapeGeometry = new ShapeGeometry(pPicture.ShapeProperties!);
    }

    public IImage Image { get; }
   
    public string? SvgContent => this.GetSvgContent();
    
    public override Geometry GeometryType
    {
        get => this.shapeGeometry.GeometryType;
        set => this.shapeGeometry.GeometryType = value;
    }

    public override decimal CornerSize
    {
        get => this.shapeGeometry.CornerSize;
        set => this.shapeGeometry.CornerSize = value;
    }

    public override decimal[] Adjustments
    {
        get => this.shapeGeometry.Adjustments;
        set => this.shapeGeometry.Adjustments = value;
    }

    public override ShapeType ShapeType => ShapeType.Picture;
    
    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }

    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removeable => true;

    public CroppingFrame Crop
    {
        get
        {
            var pic = this.pPicture;
            var aBlipFill = pic.BlipFill
                ?? throw new SCException("Malformed image has no blip fill");

            var aSrcRect = aBlipFill.GetFirstChild<A.SourceRectangle>();

            return CroppingFrameFromSourceRectangle(aSrcRect);
        }
        
        set
        {
            var pic = this.pPicture;
            var aBlipFill = pic.BlipFill
                ?? throw new SCException("Malformed image has no blip fill");

            var aSrcRect = aBlipFill.GetFirstChild<A.SourceRectangle>()
                ?? aBlipFill.InsertAfter<A.SourceRectangle>(new(), this.aBlip)
                ?? throw new SCException("Failed to add source rectangle");

            ApplyCropToSourceRectangle(value, aSrcRect);
        }
    }
    
    public decimal Transparency
    {
        get
        {
            var aAlphaModFix = this.aBlip.GetFirstChild<A.AlphaModulationFixed>();
            var amount = aAlphaModFix?.Amount?.Value ?? 100000m;
            
            return 100m - (amount / 1000m); // value is stored in Open XML as thousandths of a percent
        }

        set
        {
            var aAlphaModFix = this.aBlip.GetFirstChild<A.AlphaModulationFixed>()
                ?? this.aBlip.InsertAt<A.AlphaModulationFixed>(new(), 0)
                ?? throw new SCException("Failed to add AlphaModFix");
            
            aAlphaModFix.Amount = Convert.ToInt32((100m - value) * 1000m);
        }
    }
   
    public override void Remove() => this.pPicture.Remove();
    
    public void SendToBack()
    {
        var parentPShapeTree = this.PShapeTreeElement.Parent!;
        parentPShapeTree.RemoveChild(this.pPicture);
        var pGrpSpPr = parentPShapeTree.GetFirstChild<P.GroupShapeProperties>() !;
        pGrpSpPr.InsertAfterSelf(this.pPicture);
    }

    internal override void CopyTo(P.ShapeTree pShapeTree)
    {
        base.CopyTo(pShapeTree);

        var sourceSdkSlidePart = this.OpenXmlPart;
        var sourceImagePart = (ImagePart)sourceSdkSlidePart.GetPartById(this.blipEmbed.Value!);

        var targetImagePartRId = new SCOpenXmlPart(this.OpenXmlPart).NextRelationshipId();

        var targetImagePart = this.OpenXmlPart.AddNewPart<ImagePart>(sourceImagePart.ContentType, targetImagePartRId);
        using var sourceImageStream = sourceImagePart.GetStream(FileMode.Open);
        sourceImageStream.Position = 0;
        targetImagePart.FeedData(sourceImageStream);

        var copy = this.PShapeTreeElement.CloneNode(true);
        copy.Descendants<A.Blip>().First().Embed = targetImagePartRId;
    }

    /// <summary>
    ///     Set the cropping frame values onto the supplied source rectangle.
    /// </summary>
    /// <param name="frame">Source values to get cropping values.</param>
    /// <param name="aSrcRect">Rectangle to be updated with our values.</param>
    private static void ApplyCropToSourceRectangle(CroppingFrame frame, A.SourceRectangle aSrcRect)
    {
        aSrcRect.Left = ToThousandths(frame.Left);
        aSrcRect.Right = ToThousandths(frame.Right);
        aSrcRect.Top = ToThousandths(frame.Top);
        aSrcRect.Bottom = ToThousandths(frame.Bottom);        
    }

    /// <summary>
    ///     Convert a source rectangle to a cropping frame.
    /// </summary>
    /// <param name="aSrcRect">Source rectangle which contains the needed frame.</param>
    /// <returns>Resulting frame.</returns>
    private static CroppingFrame CroppingFrameFromSourceRectangle(A.SourceRectangle? aSrcRect)
    {
        if (aSrcRect is null)
        {
            return new CroppingFrame(0, 0, 0, 0);
        }

        return new CroppingFrame(
            FromThousandths(aSrcRect.Left),
            FromThousandths(aSrcRect.Right),
            FromThousandths(aSrcRect.Top),
            FromThousandths(aSrcRect.Bottom));
    }

    /// <summary>
    ///     Convert a value from 'percent mille' (thousandths of a percent) to percent.
    /// </summary>
    /// <param name="int32">Per cent mille value.</param>
    /// <returns>Percent value.</returns>
    private static decimal FromThousandths(Int32Value? int32) => 
        int32 is not null ? int32 / 1000m : 0;

    /// <summary>
    ///     Convert a value from 'percent mille' (thousandths of a percent).
    /// </summary>
    /// <param name="input">Percent value.</param>
    /// <returns>Per cent mille value.</returns>
    private static Int32Value? ToThousandths(decimal input) => 
        input == 0 ? null : Convert.ToInt32(input * 1000m);

    private string? GetSvgContent()
    {
        var bel = this.aBlip.GetFirstChild<A.BlipExtensionList>();
        var svgBlipList = bel?.Descendants<SVGBlip>();
        if (svgBlipList == null)
        {
            return null;
        }

        var svgId = svgBlipList.First().Embed!.Value!;

        var imagePart = (ImagePart)this.OpenXmlPart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }
}
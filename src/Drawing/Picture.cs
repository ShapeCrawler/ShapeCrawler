using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class Picture : IPicture
{
    private readonly Shape shape;
    private readonly P.Picture pPicture;
    private readonly A.Blip aBlip;
    private readonly ShapeGeometry shapeGeometry;

    internal Picture(P.Picture pPicture, A.Blip aBlip)
    {
        this.shape = new Shape(pPicture);
        this.pPicture = pPicture;
        this.aBlip = aBlip;
        this.Outline = new SlideShapeOutline(pPicture.ShapeProperties!);
        this.Fill = new ShapeFill(pPicture.ShapeProperties!);
        this.shapeGeometry = new ShapeGeometry(pPicture.ShapeProperties!);
    }

    public decimal X
    {
        get => this.shape.X;
        set => this.shape.X = value;
    }

    public decimal Y
    {
        get => this.shape.Y;
        set => this.shape.Y = value;
    }

    public IImage Image => new SlidePictureImage(this.aBlip);

    public string? SvgContent => this.GetSvgContent();

    public Geometry GeometryType
    {
        get => this.shapeGeometry.GeometryType;
        set => this.shapeGeometry.GeometryType = value;
    }

    public decimal CornerSize
    {
        get => this.shapeGeometry.CornerSize;
        set => this.shapeGeometry.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => this.shapeGeometry.Adjustments;
        set => this.shapeGeometry.Adjustments = value;
    }

    public decimal Width
    {
        get => this.shape.Width;
        set => this.shape.Width = value;
    }

    public decimal Height
    {
        get => this.shape.Height;
        set => this.shape.Height = value;
    }

    public int Id => this.shape.Id;

    public string Name
    {
        get => this.shape.Name;
        set => this.shape.Name = value;
    }

    public string AltText
    {
        get => this.shape.AltText;
        set => this.shape.AltText = value;
    }

    public bool Hidden => this.shape.Hidden;

    public PlaceholderType? PlaceholderType => this.shape.PlaceholderType;

    public string? CustomData { get; set; }

    public ShapeContent ShapeContent => ShapeContent.Picture;

    public bool HasOutline => true;

    public IShapeOutline Outline { get; }

    public bool HasFill => true;

    public IShapeFill Fill { get; }

    public ITextBox? TextBox => null;

    public double Rotation => this.shape.Rotation;

    public bool Removeable => true;

    public string SDKXPath => this.shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => this.shape.SDKOpenXmlElement;

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

    public IPresentation Presentation => this.shape.Presentation;

    public void Remove()
    {
        throw new NotImplementedException();
    }

    public ITable AsTable() => throw new SCException("Picture cannot be converted to table");

    public IMediaShape AsMedia()
    {
        throw new NotImplementedException();
    }

    public void Duplicate()
    {
        throw new NotImplementedException();
    }

    public void SendToBack()
    {
        var parentPShapeTree = this.pPicture.Parent!;
        parentPShapeTree.RemoveChild(this.pPicture);
        var pGrpSpPr = parentPShapeTree.GetFirstChild<P.GroupShapeProperties>() !;
        pGrpSpPr.InsertAfterSelf(this.pPicture);
    }

    internal void CopyTo(P.ShapeTree pShapeTree)
    {
        new SCPShapeTree(pShapeTree).Add(this.pPicture);

        var openXmlPart = this.pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var sourceSdkSlidePart = openXmlPart;
        var sourceImagePart = (ImagePart)sourceSdkSlidePart.GetPartById(this.aBlip.Embed!.Value!);

        var targetImagePartRId = new SCOpenXmlPart(openXmlPart).GetNextRelationshipId();

        var targetImagePart = openXmlPart.AddNewPart<ImagePart>(sourceImagePart.ContentType, targetImagePartRId);
        using var sourceImageStream = sourceImagePart.GetStream(FileMode.Open);
        sourceImageStream.Position = 0;
        targetImagePart.FeedData(sourceImageStream);

        var copy = this.pPicture.CloneNode(true);
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
        var openXmlPart = this.pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var bel = this.aBlip.GetFirstChild<A.BlipExtensionList>();
        var svgBlipList = bel?.Descendants<SVGBlip>();
        if (svgBlipList == null)
        {
            return null;
        }

        var svgId = svgBlipList.First().Embed!.Value!;

        var imagePart = (ImagePart)openXmlPart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }
}
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a picture shape.
/// </summary>
public interface IPicture : IShape
{
    /// <summary>
    ///     Gets image. Returns <see langword="null"/> if the content of the picture element is not a binary image. 
    /// </summary>
    IImage? Image { get; }

    /// <summary>
    ///     Gets SVG content. Returns <see langword="null"/> if the content of the picture element is not an SVG graphic.
    /// </summary>
    string? SvgContent { get; }

    /// <summary>
    ///     Gets or sets the cropping frame for this image.
    /// </summary>
    CroppingFrame Crop { get; set; }

    /// <summary>
    ///     Gets or sets the transparency for this image. Range is 0 (fully opaque, default) to 100 (fully transparent).
    /// </summary>
    decimal Transparency { get; set; }

    /// <summary>
    ///     Sends the shape backward in the z-order.
    /// </summary>
    void SendToBack();
}

internal sealed class Picture(Shape shape, P.Picture pPicture, A.Blip aBlip): IPicture
{
    public IImage Image => new SlidePictureImage(aBlip);

    public string? SvgContent => this.GetSvgContent();
    
    public ShapeContent ShapeContent => ShapeContent.Image;
    
    public bool HasOutline => true;
    
    public bool Removable => true;

    public IShapeOutline Outline => shape.Outline;

    public bool HasFill => true;

    public IShapeFill Fill => shape.Fill;

    public ITextBox? TextBox => null;
    
    public CroppingFrame Crop
    {
        get
        {
            var pic = pPicture;
            var aBlipFill = pic.BlipFill
                            ?? throw new SCException("Malformed image has no blip fill");

            var aSrcRect = aBlipFill.GetFirstChild<A.SourceRectangle>();

            return CroppingFrameFromSourceRectangle(aSrcRect);
        }

        set
        {
            var pic = pPicture;
            var aBlipFill = pic.BlipFill
                            ?? throw new SCException("Malformed image has no blip fill");

            var aSrcRect = aBlipFill.GetFirstChild<A.SourceRectangle>()
                           ?? aBlipFill.InsertAfter<A.SourceRectangle>(new(), aBlip)
                           ?? throw new SCException("Failed to add source rectangle");

            ApplyCropToSourceRectangle(value, aSrcRect);
        }
    }

    public decimal Transparency
    {
        get
        {
            var aAlphaModFix = aBlip.GetFirstChild<A.AlphaModulationFixed>();
            var amount = aAlphaModFix?.Amount?.Value ?? 100000m;

            return 100m - (amount / 1000m); // value is stored in Open XML as thousandths of a percent
        }

        set
        {
            var aAlphaModFix = aBlip.GetFirstChild<A.AlphaModulationFixed>()
                               ?? aBlip.InsertAt<A.AlphaModulationFixed>(new(), 0)
                               ?? throw new SCException("Failed to add AlphaModFix");

            aAlphaModFix.Amount = Convert.ToInt32((100m - value) * 1000m);
        }
    }
    
    public bool IsGroup => false;

    public Geometry GeometryType
    {
        get => shape.GeometryType;
        set => shape.GeometryType = value;
    }

    public decimal CornerSize
    {
        get => shape.CornerSize;
        set => shape.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => shape.Adjustments;
        set => shape.Adjustments = value;
    }

    public decimal Width
    {
        get => shape.Width;
        set => shape.Width = value;
    }

    public decimal Height
    {
        get => shape.Height;
        set => shape.Height = value;
    }
    
    public decimal X
    {
        get => shape.X;
        set => shape.X = value;
    }

    public decimal Y
    {
        get => shape.Y;
        set => shape.Y = value;
    }

    public int Id => shape.Id;

    public string Name
    {
        get => shape.Name;
        set => shape.Name = value;
    }

    public string AltText
    {
        get => shape.AltText;
        set => shape.AltText = value;
    }

    public bool Hidden => shape.Hidden;

    public PlaceholderType? PlaceholderType => shape.PlaceholderType;

    public string? CustomData
    {
        get => shape.CustomData;
        set => shape.CustomData = value;
    }

    public double Rotation => shape.Rotation;

    public string SDKXPath => shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;

    public IShapeCollection GroupedShapes => throw new SCException($"Picture is not a group. Use {nameof(IShape.ShapeContent)} property to check if the shape is a group.");

    public IPresentation Presentation => shape.Presentation;

    public void Remove()
    {
        throw new NotImplementedException();
    }

    public ITable AsTable() => throw new SCException($"Picture cannot be converted to table. Use {nameof(IShape.ShapeContent)} property to check if the shape is a table.");

    public IMediaShape AsMedia()
    {
        throw new NotImplementedException();
    }

    public void Duplicate()
    {
        throw new NotImplementedException();
    }

    public void SetText(string text)
    {
        throw new NotImplementedException();
    }

    public void SendToBack()
    {
        var parentPShapeTree = pPicture.Parent!;
        parentPShapeTree.RemoveChild(pPicture);
        var pGrpSpPr = parentPShapeTree.GetFirstChild<P.GroupShapeProperties>() !;
        pGrpSpPr.InsertAfterSelf(pPicture);
    }

    public void SetImage(string file)
    {
        using var imageStream = new FileStream(file, FileMode.Open, FileAccess.Read);
        this.Image.Update(imageStream);
    }

    internal void CopyTo(P.ShapeTree pShapeTree)
    {
        // Clone the picture and add it to the target shape tree
        new SCPShapeTree(pShapeTree).Add(pPicture);
        
        // Get the source slide part and target slide part
        var sourceOpenXmlPart = pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var sourceImagePart = (ImagePart)sourceOpenXmlPart.GetPartById(aBlip.Embed!.Value!);
        var targetOpenXmlPart = pShapeTree.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        
        // If source and target parts are the same, no need to create a new relationship
        if (sourceOpenXmlPart == targetOpenXmlPart)
        {
            return;
        }
        
        // Source and target are different slides, so we need to create a proper relationship
        // Read the source image
        using var sourceImageStream = sourceImagePart.GetStream(FileMode.Open);
        sourceImageStream.Position = 0;
        
        // Determine target part relationship ID
        string targetImagePartRId = new SCOpenXmlPart(targetOpenXmlPart).NextRelationshipId();
        
        // Create a new image part in the target slide
        var targetImagePart = targetOpenXmlPart.AddNewPart<ImagePart>(sourceImagePart.ContentType, targetImagePartRId);
        targetImagePart.FeedData(sourceImageStream);
        
        // Update the copied shape with the correct relationship ID
        var copyElement = pShapeTree.Elements<P.Picture>().Last();
        copyElement.Descendants<A.Blip>().First().Embed = targetImagePartRId;
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
        var openXmlPart = pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var bel = aBlip.GetFirstChild<A.BlipExtensionList>();
        var svgBlipList = bel?.Descendants<SVGBlip>();
        if (svgBlipList == null)
        {
            return null;
        }

        var svgId = svgBlipList.First().Embed!.Value!;

        var imagePart = (ImagePart)openXmlPart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(FileMode.Open);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }
}
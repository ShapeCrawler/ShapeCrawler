using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a shape containing video content.
/// </summary>
public interface IMediaShape : IShape
{
    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    // ReSharper disable once InconsistentNaming
    string MIME { get; }

    /// <summary>
    ///     Gets bytes of video content.
    /// </summary>
    public byte[] AsByteArray();

#if DEBUG
    /// <summary>
    ///     Gets or sets audio start mode.
    /// </summary>
    AudioStartMode StartMode { get; set; }
#endif
}

internal class MediaShape(Shape shape, SlideShapeOutline outline, ShapeFill fill, P.Picture pPicture) : IMediaShape
{
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

    public ShapeContent ShapeContent => ShapeContent.Video;

    public IShapeOutline Outline => outline;

    public IShapeFill Fill => fill;

    public ITextBox? TextBox => shape.TextBox;

    public double Rotation => shape.Rotation;

    public string SDKXPath => shape.SDKXPath;
    
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

    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;

    public IPresentation Presentation => shape.Presentation;

    public bool Removable => true;

    public string MIME
    {
        get
        {
            var openXmlPart = pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
            var p14Media = pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship =
                openXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

            return relationship.DataPart.ContentType;
        }
    }

    public AudioStartMode StartMode
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    public byte[] AsByteArray()
    {
        var openXmlPart = pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var p14Media = pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = openXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        stream.Close();

        return ms.ToArray();
    }

    public void Remove() => shape.Remove();

    public ITable AsTable() => shape.AsTable();

    public IMediaShape AsMedia() => shape.AsMedia();

    public void Duplicate() => shape.Duplicate();

    public void SetText(string text) => shape.SetText(text);

    public void SetImage(string imagePath) => shape.SetImage(imagePath);

    public void SetFontName(string fontName) => shape.SetFontName(fontName);

    public void SetFontSize(decimal fontSize) => shape.SetFontSize(fontSize);

    public void SetFontColor(string colorHex) => shape.SetFontColor(colorHex);
}
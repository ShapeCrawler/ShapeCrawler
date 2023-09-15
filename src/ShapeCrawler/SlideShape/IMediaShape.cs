using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape containing video content.
/// </summary>
public interface IMediaShape : IShape
{
    /// <summary>
    ///     Gets bytes of video content.
    /// </summary>
    public byte[] AsByteArray();

    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }
}

internal record SlideMediaShape : IMediaShape, IRemoveable
{
    private readonly SlidePart sdkSlidePart;
    private readonly P.Picture pPicture;
    private readonly SimpleShape simpleShape;

    internal SlideMediaShape(SlidePart sdkSlidePart, P.Picture pPicture)
        : this(sdkSlidePart, pPicture, new SimpleShape(pPicture))
    {
    }

    private SlideMediaShape(SlidePart sdkSlidePart, P.Picture pPicture, SimpleShape simpleShape)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pPicture = pPicture;
        this.simpleShape = simpleShape;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pPicture.ShapeProperties!);
        this.Fill = new SlideShapeFill(sdkSlidePart, pPicture.ShapeProperties!, false);
    }
    public SCShapeType ShapeType => SCShapeType.Video;
    public IMediaShape AsMedia() => this;
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public bool HasFill => true;
    public IShapeFill Fill { get; }
    
    public string MIME
    {
        get
        {
            var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship = this.sdkSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
            
            return relationship.DataPart.ContentType;
        }
    }
    
    #region SimpleShape
    
    public int X
    {
        get => this.simpleShape.X;
        set => this.simpleShape.X = value;
    }

    public int Y
    {
        get => this.simpleShape.Y;
        set => this.simpleShape.Y = value;
    }

    public int Width
    {
        get => this.simpleShape.Width;
        set => this.simpleShape.Width = value;
    }

    public int Height
    {
        get => this.simpleShape.Height;
        set => this.simpleShape.Height = value;
    }

    public int Id => this.simpleShape.Id;

    public string Name => this.simpleShape.Name;

    public bool Hidden => this.simpleShape.Hidden;

    public SCGeometry GeometryType => this.simpleShape.GeometryType;

    public bool IsPlaceholder => this.simpleShape.IsPlaceholder;

    public IPlaceholder Placeholder => this.simpleShape.Placeholder;

    public string? CustomData
    {
        get => this.simpleShape.CustomData;
        set => this.simpleShape.CustomData = value;
    }

    public bool IsTextHolder => this.simpleShape.IsTextHolder;

    public ITextFrame TextFrame => this.simpleShape.TextFrame;
    public double Rotation => this.simpleShape.Rotation;
    public ITable AsTable() => this.simpleShape.AsTable();
    
    #endregion SimpleShape
    
    public byte[] AsByteArray()
    {
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.sdkSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var bytes = stream.ToArray();
        stream.Close();

        return bytes;
    }
    
    void IRemoveable.Remove() => this.pPicture.Remove();
}
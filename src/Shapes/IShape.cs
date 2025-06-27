using System.IO;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;

// ReSharper disable InconsistentNaming
#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a slide element.
/// </summary>
public interface IShape : IPosition, IShapeGeometry
{
    /// <summary>
    ///     Gets or sets the width of the shape in points.
    /// </summary>
    decimal Width { get; set; }

    /// <summary>
    ///     Gets or sets the height of the shape in points.
    /// </summary>
    decimal Height { get; set; }

    /// <summary>
    ///     Gets identifier of the shape.
    /// </summary>
    int Id { get; }
    
    /// <summary>
    ///    Gets or sets the name of the shape.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    ///     Gets or sets the shape alternative text.
    /// </summary>
    string AltText { get; set; }

    /// <summary>
    ///     Gets a value indicating whether the shape is hidden.
    /// </summary>
    bool Hidden { get; }
    
    /// <summary>
    ///     Gets the placeholder type. Returns <see langword="null"/> if the shape is not a placeholder.
    /// </summary>
    PlaceholderType? PlaceholderType { get; }

    /// <summary>
    ///     Gets or sets custom data string for the shape.
    /// </summary>
    string? CustomData { get; set; }
    
    /// <summary>
    ///     Gets outline of the shape. Returns <see langword="null"/> if the shape cannot be outlined, for example, a picture.
    /// </summary>
    IShapeOutline? Outline { get; }
    
    /// <summary>
    ///     Gets the fill of the shape. Returns <see langword="null"/> if the shape cannot be filled, for example, a line.
    /// </summary>
    IShapeFill? Fill { get; }
    
    /// <summary>
    ///     Gets Text Box. Returns <c>null</c> if the shape is not a text holder.
    /// </summary>
    ITextBox? TextBox { get; }
    
    /// <summary>
    ///     Gets picture. Returns <c>null</c> if the shape doesn't contain image content.
    /// </summary>
    IPicture? Picture { get; }
    
    /// <summary>
    ///     Gets chart. Returns <c>null</c> if the shape doesn't contain image content.
    /// </summary>
    IChart? Chart { get; }
    
    /// <summary>
    ///     Gets table. Returns <c>null</c> if the shape doesn't contain table content.
    /// </summary>
    ITable? Table { get; }
    
    /// <summary>
    ///     Gets OLE object. Returns <c>null</c> if the shape doesn't contain OLE object content.
    /// </summary>
    IOLEObject? OLEObject { get; }
    
    /// <summary>
    ///     Gets media. Returns <c>null</c> if the shape doesn't contain media content.
    /// </summary>
    IMedia? Media { get; }
    
    /// <summary>
    ///     Gets line. Returns <c>null</c> if the shape is not a line.
    /// </summary>
    ILine? Line { get; }
    
    /// <summary>
    ///     Gets the rotation of the shape in degrees.
    /// </summary>
    double Rotation { get; }
    
    /// <summary>
    ///     Gets XPath of the underlying Open XML element.
    /// </summary>
    public string SDKXPath { get; }
    
    /// <summary>
    ///     Gets a copy of the underlying Open XML element.
    /// </summary>
    OpenXmlElement SDKOpenXmlElement { get; }
    
    /// <summary>
    ///     Gets the parent presentation.
    /// </summary>
    IPresentation Presentation { get; }

    /// <summary>
    ///     Gets grouped shapes. Returns <c>null</c> if the shape is not a group shape.
    /// </summary>
    IShapeCollection? GroupedShapes { get; }

    /// <summary>
    ///     Removes the shape from the slide.
    /// </summary>
    void Remove();

    /// <summary>
    ///     Duplicates the shape.
    /// </summary>
    void Duplicate();

    /// <summary>
    ///     Sets text content. 
    ///     Use property <see cref="TextBox"/> to check whether the shape is a text holder.
    /// </summary>
    /// <exception cref="SCException">Thrown when the shape is not a text holder.</exception>
    void SetText(string text);
    
    /// <summary>
    ///     Sets text content. 
    ///     Use property <see cref="TextBox"/> to check whether the shape is a text holder.
    /// </summary>
    /// <exception cref="SCException">Thrown when the shape is not a text holder.</exception>
    void SetMarkdownText(string text);

    /// <summary>
    ///     Sets image content.
    ///     Use <see cref="Picture"/> property to check whether shape contains image content.
    /// </summary>
    /// <exception cref="SCException">Thrown if the shape doesn't contain image content.</exception>
    void SetImage(string imagePath);
    
    /// <summary>
    ///     Sets the font name.
    /// </summary>
    void SetFontName(string fontName);
    
    /// <summary>
    ///     Sets the font size.
    /// </summary>
    void SetFontSize(decimal fontSize);
    
    /// <summary>
    ///     Sets the font color.
    /// </summary>
    void SetFontColor(string colorHex);

    /// <summary>
    ///     Sets the video content.
    /// </summary>
    /// <exception cref="SCException">Thrown if the shape is not video content holder.</exception>
    void SetVideo(Stream video);
}
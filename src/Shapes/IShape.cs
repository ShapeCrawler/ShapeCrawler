using System.IO;
using DocumentFormat.OpenXml;

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
    ///     Gets the shape content type.
    /// </summary>
    ShapeContent ShapeContent { get; }
    
    /// <summary>
    ///     Gets outline of the shape. Returns <see langword="null"/> if the shape cannot be outlined, for example, a picture.
    /// </summary>
    IShapeOutline? Outline { get; }
    
    /// <summary>
    ///     Gets the fill of the shape. Returns <see langword="null"/> if the shape cannot be filled, for example, a line.
    /// </summary>
    IShapeFill? Fill { get; }

    /// <summary>
    ///     Gets Text Box. Returns <c>null</c> if the slide element doesn't contain text content. Use <see cref="ShapeContent"/> property to check content type.
    /// </summary>
    ITextBox? TextBox { get; }
    
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
    ///     Removes the shape from the slide.
    /// </summary>
    void Remove();
    
    /// <summary>
    ///     Gets the table if the shape is a table. Use <see cref="ShapeContent"/> property to check if the shape is a table.
    /// </summary>
    ITable AsTable();
    
    /// <summary>
    ///     Gets the media shape which is an audio or video.
    ///     Use <see cref="ShapeContent"/> property to check if the shape is an audio or video.    
    /// </summary>
    IMediaShape AsMedia();

    /// <summary>
    ///     Duplicates the shape.
    /// </summary>
    void Duplicate();

    /// <summary>
    ///     Sets the text content. Throws <see cref="SCException"/> if text content cannot be set for this element.
    ///     Use <see cref="ShapeContent"/> property to check element content type.
    /// </summary>
    void SetText(string text);

    /// <summary>
    ///     Sets the image content. Throws <see cref="SCException"/> if image content cannot be set for this element.
    ///     Use <see cref="ShapeContent"/> property to check element content type.
    /// </summary>
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
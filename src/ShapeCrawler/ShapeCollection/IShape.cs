using DocumentFormat.OpenXml;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a shape.
/// </summary>
public interface IShape : IPosition, IShapeGeometry
{
    /// <summary>
    ///     Gets or sets the width of the shape in pixels.
    /// </summary>
    decimal Width { get; set; }

    /// <summary>
    ///     Gets or sets the height of the shape in pixels.
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
    ///     Gets or sets the alternative text for the shape.
    /// </summary>
    string AltText { get; set; }

    /// <summary>
    ///     Gets a value indicating whether the shape is hidden.
    /// </summary>
    bool Hidden { get; }

    /// <summary>
    ///     Gets a value indicating whether shape is a placeholder.
    /// </summary>
    bool IsPlaceholder { get; }
    
    /// <summary>
    ///     Gets the placeholder type of the shape.
    /// </summary>
    PlaceholderType PlaceholderType { get; }

    /// <summary>
    ///     Gets or sets custom data string for the shape.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets the type of shape.
    /// </summary>
    ShapeType ShapeType { get; }
    
    /// <summary>
    ///     Gets a value indicating whether the shape has outline formatting.
    /// </summary>
    bool HasOutline { get; }
    
    /// <summary>
    ///     Gets outline of the shape.
    /// </summary>
    IShapeOutline Outline { get; }
 
    /// <summary>
    ///     Gets a value indicating whether the shape has fill.
    /// </summary>
    bool HasFill { get; }
    
    /// <summary>
    ///     Gets the fill of the shape. Returns <see langword="null"/> if the shape cannot be filled, for example, a line.
    /// </summary>
    IShapeFill Fill { get; }
    
    /// <summary>
    ///     Gets a value indicating whether the shape is a text holder.
    /// </summary>
    bool IsTextHolder { get; }
    
    /// <summary>
    ///     Gets Text Frame.
    /// </summary>
    ITextBox TextBox { get; }
    
    /// <summary>
    ///     Gets the rotation of the shape in degrees.
    /// </summary>
    double Rotation { get; }
    
    /// <summary>
    ///     Gets a value indicating whether the shape can be removed.
    /// </summary>
    bool Removeable { get; }
    
    /// <summary>
    ///     Gets XPath of the underlying Open XML element.
    /// </summary>
    public string SdkXPath { get; }
    
    /// <summary>
    ///     Gets a copy of the underlying Open XML element.
    /// </summary>
    OpenXmlElement SdkOpenXmlElement { get; }

    /// <summary>
    ///     Gets or sets the text content of the shape.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    ///     Removes the shape from the slide.
    /// </summary>
    void Remove();
    
    /// <summary>
    ///     Gets the table if the shape is a table. Use <see cref="IShape.ShapeType"/> property to check if the shape is a table.
    /// </summary>
    ITable AsTable();
    
    /// <summary>
    ///     Gets the media shape which is an audio or video.
    ///     Use <see cref="IShape.ShapeType"/> property to check if the shape is an audio or video.    
    /// </summary>
    IMediaShape AsMedia();
}
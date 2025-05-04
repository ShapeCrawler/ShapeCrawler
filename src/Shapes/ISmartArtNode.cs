namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a node in a SmartArt graphic.
/// </summary>
public interface ISmartArtNode
{
    /// <summary>
    ///     Gets or sets the text of the SmartArt node.
    /// </summary>
    string Text { get; set; }
}

/// <summary>
///     Represents a node in a SmartArt graphic.
/// </summary>
internal class SmartArtNode : ISmartArtNode
{
    private readonly SmartArtNodeCollection nodeCollection;
    private string textValue;

    internal SmartArtNode(string modelId, string text, SmartArtNodeCollection nodeCollection)
    {
        this.ModelId = modelId;
        this.textValue = text;
        this.nodeCollection = nodeCollection;
    }
    
    /// <summary>
    ///     Gets or sets the text of the SmartArt node.
    /// </summary>
    public string Text
    {
        get => this.textValue;
        set
        {
            if (this.textValue != value)
            {
                this.textValue = value;
                this.nodeCollection?.UpdateNodeText(this.ModelId, value);
            }
        }
    }
    
    internal string ModelId { get; }
    
    internal void UpdateText(string text)
    {
        this.textValue = text;
    }
}
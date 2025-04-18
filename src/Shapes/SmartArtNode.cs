using System;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a node in a SmartArt graphic.
/// </summary>
internal class SmartArtNode : ISmartArtNode
{
    private readonly SmartArtNodeCollection nodeCollection;
    private string textValue;

    internal string ModelId { get; }

    internal SmartArtNode(string modelId, string text, SmartArtNodeCollection nodeCollection)
    {
        this.ModelId = modelId ?? throw new ArgumentNullException(nameof(modelId));
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
    
    internal void UpdateText(string text)
    {
        this.textValue = text;
    }
}

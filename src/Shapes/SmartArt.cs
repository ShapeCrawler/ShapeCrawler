using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a SmartArt graphic shape.
/// </summary>
internal class SmartArt : Shape, ISmartArt
{
    private readonly OpenXmlPart dataPart;
    
    internal SmartArt(OpenXmlElement pShapeTreeElement, OpenXmlPart dataPart) 
        : base(pShapeTreeElement)
    {
        this.dataPart = dataPart; // Can be null in our simplified implementation
        this.Nodes = new SmartArtNodeCollection();
    }
    
    /// <summary>
    ///     Gets the collection of nodes in the SmartArt graphic.
    /// </summary>
    public ISmartArtNodeCollection Nodes { get; }
}

using DocumentFormat.OpenXml;

namespace ShapeCrawler.Shapes;

internal class SmartArt : Shape, ISmartArt
{
    internal SmartArt(OpenXmlElement pShapeTreeElement) 
        : base(pShapeTreeElement)
    {
        this.Nodes = new SmartArtNodeCollection();
    }
    
    public ISmartArtNodeCollection Nodes { get; }
}

using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

internal abstract class SlideStructure : ISlideStructure
{
    protected SlideStructure(PresentationCore pres)
    {
        this.PresCore = pres;
    }

    public PresentationCore PresCore { get; protected init; }
    
    public abstract int Number { get; set; }
    
    public abstract IShapeCollection Shapes { get; }

    internal abstract TypedOpenXmlPart TypedOpenXmlPart { get; }

    internal int GetNextShapeId()
    {
        if (this.Shapes.Any())
        {
           return this.Shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;    
        }

        return 1;
    }
}
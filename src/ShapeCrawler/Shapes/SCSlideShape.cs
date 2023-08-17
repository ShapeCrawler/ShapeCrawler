using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a slide shape.
/// </summary>
public abstract record SCSlideShape
{
    private readonly TypedOpenXmlCompositeElement pShapeTreeChild;

    protected internal SCSlideShape(TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.pShapeTreeChild = pShapeTreeChild;
    }
    
    internal abstract void CopyTo(int id, ShapeTree targetPShapeTree, IEnumerable<string> targetExistingShapeNames, SlidePart targetSdkSlidePart);
}
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a slide shape.
/// </summary>
public abstract record SCSlideShape
{
    private TypedOpenXmlCompositeElement pShapeTreeChild;

    protected SCSlideShape(TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.pShapeTreeChild = pShapeTreeChild;
    }

    internal abstract bool Copyable();

    internal OpenXmlElement CopyUnderlyingTypedOpenXmlCompositeElement()
    {
        return this.pShapeTreeChild.CloneNode(true);
    }
}
using DocumentFormat.OpenXml;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Provides APIs to create shape tree's elements.
    /// </summary>
    public interface IElementFactory
    {
        ShapeEx ElementFromCandidate(ElementCandidate ec, IPreSettings parents);

        ShapeEx GroupFromXml(OpenXmlCompositeElement compositeElement, IPreSettings parents);
    }
}
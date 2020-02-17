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
        Shape ElementFromCandidate(ElementCandidate ec, IParents parents);

        Shape GroupFromXml(OpenXmlCompositeElement compositeElement, IParents parents);
    }
}
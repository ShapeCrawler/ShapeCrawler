using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents;

namespace SlideXML.Services
{
    /// <summary>
    /// Provides APIs to create shape tree's elements.
    /// </summary>
    public interface IElementFactory
    {
        SlideElement CreateShape(ElementCandidate ec, IPreSettings preSettings);

        SlideElement CreateGroupShape(OpenXmlCompositeElement compositeElement, IPreSettings preSettings);
    }
}
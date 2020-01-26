using DocumentFormat.OpenXml;

namespace SlideXML.Services.Placeholders
{
    /// <summary>
    /// Provides APIs for placeholder service.
    /// </summary>
    public interface IPlaceholderService
    {
        PlaceholderSL Get(OpenXmlCompositeElement ce);
    }
}
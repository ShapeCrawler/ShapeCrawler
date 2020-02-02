using DocumentFormat.OpenXml;

namespace SlideXML.Services.Placeholders
{
    /// <summary>
    /// Represents a Slide Layout placeholder service.
    /// </summary>
    public interface IPlaceholderService
    {
        PlaceholderSL TryGet(OpenXmlCompositeElement ce);
    }
}
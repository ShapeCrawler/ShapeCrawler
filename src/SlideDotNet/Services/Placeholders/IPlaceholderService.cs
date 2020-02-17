using DocumentFormat.OpenXml;

namespace SlideDotNet.Services.Placeholders
{
    /// <summary>
    /// Represents a Slide Layout placeholder service.
    /// </summary>
    public interface IPlaceholderService
    {
        PlaceholderLocationData TryGet(OpenXmlCompositeElement ce);
    }
}
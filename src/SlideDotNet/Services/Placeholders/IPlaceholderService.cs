using DocumentFormat.OpenXml;

namespace SlideDotNet.Services.Placeholders
{
    /// <summary>
    /// Represents a Slide Layout placeholder service.
    /// </summary>
    public interface IPlaceholderService
    {
        /// <summary>
        /// Gets placeholder from the repository. Returns null if data does not exist for specified element.
        /// </summary>
        /// <remarks>
        /// Placeholder can have their location and size property values data on the slide.
        /// </remarks>
        PlaceholderLocationData TryGet(OpenXmlCompositeElement sdkElement);
    }
}
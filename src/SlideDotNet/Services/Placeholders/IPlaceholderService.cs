using DocumentFormat.OpenXml;
using SlideDotNet.Enums;

namespace SlideDotNet.Services.Placeholders
{
    /// <summary>
    /// Represents a Slide Layout placeholder service.
    /// </summary>
    public interface IPlaceholderService
    {
        /// <summary>
        /// Tries to get matched <see cref="PlaceholderLocationData"/> instance for specified SDK-element.
        /// Returns null if matched object is not found.
        /// </summary>
        /// <remarks>
        /// Placeholder can have their location and size property values data on the slide.
        /// </remarks>
        PlaceholderLocationData TryGetLocation(OpenXmlCompositeElement sdkCompositeElement);

        PlaceholderFontData PlaceholderFontDataFromCompositeElement(OpenXmlCompositeElement sdkCompositeElement);

        PlaceholderData CreatePlaceholderData(OpenXmlElement sdkElement);

        PlaceholderType GetPlaceholderType(OpenXmlElement sdkElement);
    }
}
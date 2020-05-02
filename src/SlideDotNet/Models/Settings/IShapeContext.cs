using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Services;
using SlideDotNet.Services.Placeholders;

namespace SlideDotNet.Models.Settings
{
    /// <summary>
    /// Represents a shape context.
    /// </summary>
    public interface IShapeContext
    {
        /// <summary>
        /// Returns a presentation settings.
        /// </summary>
        IPreSettings PreSettings { get; }

        /// <summary>
        /// Returns a service for placeholder's fonts.
        /// </summary>
        PlaceholderFontService PlaceholderFontService { get; }

        public IPlaceholderService PlaceholderService { get; }

        /// <summary>
        /// Returns a <see cref="OpenXmlElement"/> instance.
        /// </summary>
        OpenXmlElement SdkElement { get; }

        SlidePart SkdSlidePart { get; }


        bool TryGetFontHeight(int prLvl, out int fh);
    }
}
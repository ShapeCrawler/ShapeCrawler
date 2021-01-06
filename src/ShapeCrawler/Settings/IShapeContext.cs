using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories.Placeholders;

namespace ShapeCrawler.Settings
{
    /// <summary>
    /// Represents a shape context.
    /// </summary>
    public interface IShapeContext
    {
        /// <summary>
        /// Returns a presentation settings.
        /// </summary>
        IPresentationData presentationData { get; }

        /// <summary>
        /// Returns a service for placeholder's fonts.
        /// </summary>
        PlaceholderFontService PlaceholderFontService { get; }

        public IPlaceholderService PlaceholderService { get; }

        /// <summary>
        /// Returns a <see cref="OpenXmlElement"/> instance.
        /// </summary>
        OpenXmlElement SdkElement { get; }

        SlidePart SdkSlidePart { get; }

        bool TryGetFontHeight(int prLvl, out int fh);
    }
}
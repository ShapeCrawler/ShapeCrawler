using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Services;

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
        SlidePlaceholderFontService PlaceholderFontService { get; }

        /// <summary>
        /// Returns a <see cref="OpenXmlElement"/> instance.
        /// </summary>
        OpenXmlElement SdkElement { get; }

        SlidePart SkdSlidePart { get; }


        bool TryFromMasterOther(int prLvl, out int fh);
    }
}
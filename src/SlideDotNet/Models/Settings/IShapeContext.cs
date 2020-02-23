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
        /// Returns presentation settings.
        /// </summary>
        public IPreSettings PreSettings { get; }

        public SlidePlaceholderFontService PlaceholderFontService { get; }

        public OpenXmlCompositeElement XmlElement { get; }

        /// <summary>
        /// Returns placeholder location data.
        /// </summary>
        public PlaceholderLocationData PlaceholderLocationData { get; }

        public SlidePart XmlSlidePart { get; }

        bool TryFromMasterOther(int prLvl, out int fh);
    }
}
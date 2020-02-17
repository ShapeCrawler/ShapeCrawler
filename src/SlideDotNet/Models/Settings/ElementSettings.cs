using System;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services;
using SlideDotNet.Services.Placeholders;

namespace SlideDotNet.Models.Settings
{
    /// <summary>
    /// Represents an element's settings.
    /// </summary>
    public class ElementSettings : IShapeSettings
    {
        #region Properties

        /// <summary>
        /// Returns presentation settings.
        /// </summary>
        public IParents Parents { get; }

        public SlidePlaceholderFontService FontService { get; } //TODO: consider the possibility to remove setter

        public OpenXmlCompositeElement XmlElement { get; set; }

        /// <summary>
        /// Returns placeholder data.
        /// </summary>
        public PlaceholderLocationData Placeholder { get; set; } //TODO: consider the possibility to remove setter

        public Shape SlideElement { get; set; }

        #endregion Properties

        #region Constructors

        public ElementSettings(IParents parents, SlidePlaceholderFontService fontService, OpenXmlCompositeElement xmlElement)
        {
            Parents = parents ?? throw new ArgumentNullException(nameof(parents));
            FontService = fontService ?? throw new ArgumentNullException(nameof(fontService));
            XmlElement = xmlElement ?? throw new ArgumentNullException(nameof(xmlElement));
        }

        #endregion Constructors
    }
}

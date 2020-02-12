using SlideXML.Models.SlideComponents;
using SlideXML.Services.Placeholders;

namespace SlideXML.Models.Settings
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
        public IPreSettings PreSettings { get; }

        /// <summary>
        /// Returns placeholder data.
        /// </summary>
        public PlaceholderData Placeholder { get; set; } //TODO: consider the possibility to remove setter

        public SlideElement Shape { get; set; }

        #endregion Properties

        #region Constructors

        public ElementSettings(IPreSettings preSettings)
        {
            PreSettings = preSettings;
        }

        #endregion Constructors
    }
}

using SlideXML.Services.Placeholders;

namespace SlideXML.Models.Settings
{
    /// <summary>
    /// Represent an element's settings.
    /// </summary>
    public class ElementSettings : IShapeSettings
    {
        public IPreSettings PreSettings { get; }

        public PlaceholderEx Placeholder { get; set; }

        public ElementSettings(IPreSettings preSettings)
        {
            PreSettings = preSettings;
        }
    }
}

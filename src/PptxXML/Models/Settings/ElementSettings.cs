using PptxXML.Services.Placeholders;

namespace PptxXML.Models.Settings
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

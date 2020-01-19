using SlideXML.Services.Placeholders;

namespace SlideXML.Models.Settings
{
    public interface IShapeSettings
    {
        public IPreSettings PreSettings { get; }

        public PlaceholderEx Placeholder { get; set; }
    }
}
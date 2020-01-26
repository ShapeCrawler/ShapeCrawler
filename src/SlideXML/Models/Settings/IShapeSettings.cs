using SlideXML.Services.Placeholders;

namespace SlideXML.Models.Settings
{
    public interface IShapeSettings
    {
        public IPreSettings PreSettings { get; }

        public PlaceholderSL Placeholder { get; set; }
    }
}
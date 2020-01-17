using PptxXML.Services.Placeholders;

namespace PptxXML.Models.Settings
{
    public interface IShapeSettings
    {
        public IPreSettings PreSettings { get; }

        public PlaceholderEx Placeholder { get; set; }
    }
}
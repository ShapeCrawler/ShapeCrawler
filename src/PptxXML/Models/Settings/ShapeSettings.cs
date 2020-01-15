using PptxXML.Services.Placeholders;

namespace PptxXML.Models.Settings
{
    public interface IShapeSettings
    {
        public IPreSettings PreSettings { get; }

        public Placeholder Placeholder { get; set; }
    }

    public class ShapeSettings : IShapeSettings
    {
        public IPreSettings PreSettings { get; }

        public Placeholder Placeholder { get; set; }

        public ShapeSettings(IPreSettings preSettings)
        {
            PreSettings = preSettings;
        }
    }
}

using SlideDotNet.Services.Placeholders;

namespace SlideDotNet.Models.Settings
{
    public interface IShapeSettings
    {
        public IParents Parents { get; }

        public PlaceholderLocationData Placeholder { get; set; }
    }
}
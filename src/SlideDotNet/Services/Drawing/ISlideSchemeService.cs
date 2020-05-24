using SlideDotNet.Models;

namespace SlideDotNet.Services.Drawing
{
    public interface ISlideSchemeService
    {
        void SaveScheme(string filePath, ShapeCollection shapesValue, int sldW, int sldH);
    }
}
using System.IO;
using SlideDotNet.Models;

namespace SlideDotNet.Services.Drawing
{
    public interface ISlideSchemeService
    {
        void SaveScheme(ShapeCollection shapes, int sldW, int sldH, string filePath);
        
        void SaveScheme(ShapeCollection shapes, int sldW, int sldH, Stream stream);
    }
}
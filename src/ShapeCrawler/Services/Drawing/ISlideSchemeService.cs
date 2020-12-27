using System.IO;
using ShapeCrawler.Collections;
using ShapeCrawler.Models;
using SlideDotNet.Models;

namespace ShapeCrawler.Services.Drawing
{
    public interface ISlideSchemeService
    {
        void SaveScheme(ShapeCollection shapes, int sldW, int sldH, string filePath);
        
        void SaveScheme(ShapeCollection shapes, int sldW, int sldH, Stream stream);
    }
}
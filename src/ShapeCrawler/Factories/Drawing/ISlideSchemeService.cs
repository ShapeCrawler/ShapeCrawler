using System.IO;
using ShapeCrawler.Collections;

namespace ShapeCrawler.Factories.Drawing
{
    public interface ISlideSchemeService
    {
        void SaveScheme(ShapeCollection shapes, int sldW, int sldH, string filePath);
        
        void SaveScheme(ShapeCollection shapes, int sldW, int sldH, Stream stream);
    }
}
using ShapeCrawler.Placeholders;

namespace ShapeCrawler.AutoShapes
{
    internal interface IFontDataReader
    {
        bool TryGetFontData(int paragraphLvl, out FontData fontData);
    }
}
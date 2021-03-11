using ShapeCrawler.Placeholders;

namespace ShapeCrawler.AutoShapes
{
    internal interface IAutoShapeInternal
    {
        bool TryGetFontData(int paragraphLvl, out FontData fontData);
    }
}
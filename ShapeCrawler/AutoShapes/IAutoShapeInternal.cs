using ShapeCrawler.Placeholders;

namespace ShapeCrawler.AutoShapes
{
    internal interface IAutoShapeInternal
    {
        bool TryGetFontSize(int paragraphLvl, out int i);
        bool TryGetFontData(int paragraphLvl, out FontData fontData);
    }
}
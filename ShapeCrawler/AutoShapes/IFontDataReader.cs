using ShapeCrawler.Placeholders;

namespace ShapeCrawler.AutoShapes
{
    internal interface IFontDataReader
    {
        void FillFontData(int paragraphLvl, ref FontData fontData);
    }
}
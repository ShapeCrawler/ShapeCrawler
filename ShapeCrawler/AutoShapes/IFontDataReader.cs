using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;

namespace ShapeCrawler.AutoShapes
{
    internal interface IFontDataReader
    {
        void FillFontData(int paragraphLvl, ref FontData fontData);
    }
}
using System.Diagnostics.CodeAnalysis;
using ShapeCrawler.Services;

namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "It is an internal member.")]
    internal interface IFontDataReader
    {
        void FillFontData(int paragraphLvl, ref FontData fontData);
    }
}
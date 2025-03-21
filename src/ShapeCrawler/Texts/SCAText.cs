using System.Linq;
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once InconsistentNaming
internal sealed class SCAText(A.Text aText)
{
    internal string EastAsianName()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aEastAsianFont = aText.Parent!.GetFirstChild<A.RunProperties>()?.GetFirstChild<A.EastAsianFont>();
        if (aEastAsianFont != null)
        {
            if (aEastAsianFont.Typeface == "+mj-ea")
            {
                var themeFontScheme = new ThemeFontScheme(openXmlPart);
                return themeFontScheme.MajorEastAsianFont();
            }

            return aEastAsianFont.Typeface!;
        }
        
        return new ThemeFontScheme(openXmlPart).MinorEastAsianFont();
    }

    internal void UpdateEastAsianName(string eastAsianFont)
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aEastAsianFont = aText.Parent!.GetFirstChild<A.RunProperties>()?.GetFirstChild<A.EastAsianFont>();
        if (aEastAsianFont != null)
        {
            aEastAsianFont.Typeface = eastAsianFont;
            return;
        }
        
        new ThemeFontScheme(openXmlPart).UpdateMinorEastAsianFont(eastAsianFont);
    }
}
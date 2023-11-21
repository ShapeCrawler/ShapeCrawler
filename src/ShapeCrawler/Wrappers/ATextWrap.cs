using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;

internal sealed class ATextWrap
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Text aText;

    internal ATextWrap(TypedOpenXmlPart sdkTypedOpenXmlPart, A.Text aText)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aText = aText;
    }

    internal string EastAsianName()
    {
        var aEastAsianFont = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.GetFirstChild<A.EastAsianFont>();
        if (aEastAsianFont != null)
        {
            if (aEastAsianFont.Typeface == "+mj-ea")
            {
                var themeFontScheme = new ThemeFontScheme(this.sdkTypedOpenXmlPart);
                return themeFontScheme.MajorEastAsianFont();
            }

            return aEastAsianFont.Typeface!;
        }
        
        return new ThemeFontScheme(this.sdkTypedOpenXmlPart).MinorEastAsianFont();
    }

    internal void UpdateEastAsianName(string eastAsianFont)
    {
        var aEastAsianFont = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.GetFirstChild<A.EastAsianFont>();
        if (aEastAsianFont != null)
        {
            aEastAsianFont.Typeface = eastAsianFont;
            return;
        }
        
        new ThemeFontScheme(this.sdkTypedOpenXmlPart).UpdateMinorEastAsianFont(eastAsianFont);
    }
}
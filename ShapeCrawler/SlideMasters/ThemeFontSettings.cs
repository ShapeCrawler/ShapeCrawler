using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.SlideMasters;

internal class ThemeFontSettings : IThemeFontSetting
{
    private readonly FontScheme aFontScheme;

    internal ThemeFontSettings(FontScheme aFontScheme)
    {
        this.aFontScheme = aFontScheme;
    }

    public string Head
    {
        get => this.GetHeadingFontName();
        set => this.SetHeadingFontName(value);
    }

    public string Body
    {
        get => this.GetBodyFontName();
        set => this.SetBodyFontName(value);
    }

    private string GetBodyFontName()
    {
        return this.aFontScheme.MinorFont!.LatinFont!.Typeface!.Value!;
    }
    
    private void SetHeadingFontName(string fontName)
    {
        this.aFontScheme.MajorFont!.LatinFont!.Typeface!.Value = fontName;
    }

    private void SetBodyFontName(string fontName)
    {
        this.aFontScheme.MinorFont!.LatinFont!.Typeface!.Value = fontName;
    }
    
    private string GetHeadingFontName()
    {
        return this.aFontScheme.MajorFont!.LatinFont!.Typeface!.Value!;
    }
}
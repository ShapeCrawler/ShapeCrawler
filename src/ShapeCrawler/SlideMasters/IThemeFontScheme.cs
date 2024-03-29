using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a settings of theme font.
/// </summary>
public interface IThemeFontScheme
{
    /// <summary>
    ///     Gets or sets font name for head.
    /// </summary>
    string HeadLatinFont { get; set; }

    /// <summary>
    ///     Gets or sets font name for the Latin characters of the body.
    /// </summary>
    string BodyLatinFont { get; set; }

    /// <summary>
    ///     Gets or sets font name for the East Asian characters of the body.
    /// </summary>
    string BodyEastAsianFont { get; set; }

    /// <summary>
    ///     Gets or sets font name for the East Asian characters of the heading.
    /// </summary>
    string HeadEastAsianFont { get; set; }
}

internal sealed class ThemeFontScheme : IThemeFontScheme
{
    private readonly A.FontScheme aFontScheme;

    internal ThemeFontScheme(OpenXmlPart sdkTypedOpenXmlPart)
    {
        this.aFontScheme = sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .FontScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .FontScheme!,
            _ => ((SlideMasterPart)sdkTypedOpenXmlPart).ThemePart!.Theme.ThemeElements!.FontScheme!
        };
    }

    public string HeadLatinFont
    {
        get => this.GetHeadLatinFont();
        set => this.SetHeadLatinFont(value);
    }

    public string BodyLatinFont
    {
        get => this.GetBodyLatinFont();
        set => this.SetBodyLatinFont(value);
    }

    public string BodyEastAsianFont
    {
        get => this.GetBodyEastAsianFont();
        set => this.SetBodyEastAsianFont(value);
    }

    public string HeadEastAsianFont
    {
        get => this.GetHeadEastAsianFont();
        set => this.SetHeadEastAsianFont(value);
    }

    internal string MajorLatinFont() => this.aFontScheme.MajorFont!.LatinFont!.Typeface!;

    internal string MajorEastAsianFont() => this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!;

    internal string MinorEastAsianFont() => this.aFontScheme.MinorFont!.EastAsianFont!.Typeface!;

    internal A.LatinFont MinorLatinFont() => this.aFontScheme.MinorFont!.LatinFont!;

    internal void UpdateMinorEastAsianFont(string eastAsianFont) =>
        this.aFontScheme.MinorFont!.EastAsianFont!.Typeface = eastAsianFont;

    private string GetHeadLatinFont() => this.aFontScheme.MajorFont!.LatinFont!.Typeface!.Value!;

    private string GetHeadEastAsianFont() => this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value!;

    private void SetHeadLatinFont(string fontName) => this.aFontScheme.MajorFont!.LatinFont!.Typeface!.Value = fontName;

    private void SetHeadEastAsianFont(string fontName) =>
        this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value = fontName;

    private string GetBodyLatinFont() => this.aFontScheme.MinorFont!.LatinFont!.Typeface!.Value!;

    private string GetBodyEastAsianFont() => this.aFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value!;

    private void SetBodyLatinFont(string fontName) => this.aFontScheme.MinorFont!.LatinFont!.Typeface!.Value = fontName;

    private void SetBodyEastAsianFont(string fontName) =>
        this.aFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value = fontName;
}
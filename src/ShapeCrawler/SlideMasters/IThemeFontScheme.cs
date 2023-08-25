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

    internal ThemeFontScheme(A.FontScheme aFontScheme)
    {
        this.aFontScheme = aFontScheme;
    }

    public string HeadLatinFont
    {
        get => this.GetHeadLatinFont();
        set => this.SetHeadLatinFont(value);
    }

    public string HeadEastAsianFont
    {
        get => this.GetHeadEastAsianFont();
        set => this.SetHeadEastAsianFont(value);
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

    private string GetHeadLatinFont()
    {
        return this.aFontScheme.MajorFont!.LatinFont!.Typeface!.Value!;
    }

    private string GetHeadEastAsianFont()
    {
        return this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value!;
    }

    private void SetHeadLatinFont(string fontName)
    {
        this.aFontScheme.MajorFont!.LatinFont!.Typeface!.Value = fontName;
    }

    private void SetHeadEastAsianFont(string fontName)
    {
        this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value = fontName;
    }

    private string GetBodyLatinFont()
    {
        return this.aFontScheme.MinorFont!.LatinFont!.Typeface!.Value!;
    }

    private string GetBodyEastAsianFont()
    {
        return this.aFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value!;
    }

    private void SetBodyLatinFont(string fontName)
    {
        this.aFontScheme.MinorFont!.LatinFont!.Typeface!.Value = fontName;
    }

    private void SetBodyEastAsianFont(string fontName)
    {
        this.aFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value = fontName;
    }

    public string MajorFontLatinFont()
    {
        return this.aFontScheme.MajorFont!.LatinFont!.Typeface!;
    }

    public string MajorFontEastAsianFont()
    {
        return this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!;
    }
}
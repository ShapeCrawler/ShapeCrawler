using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler;

/// <summary>
///     Represents a settings of theme font.
/// </summary>
public interface IThemeFontScheme
{
    /// <summary>
    ///     Gets or sets font name for head.
    /// </summary>
    string Head { get; set; }

    /// <summary>
    ///     Gets or sets font name for body.
    /// </summary>
    string Body { get; set; }
}

internal class ThemeFontScheme : IThemeFontScheme
{
    private readonly A.FontScheme aFontScheme;

    internal ThemeFontScheme(A.FontScheme aFontScheme)
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
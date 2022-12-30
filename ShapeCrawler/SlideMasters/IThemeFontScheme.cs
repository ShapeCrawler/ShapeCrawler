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

internal sealed class ThemeFontScheme : IThemeFontScheme
{
    internal ThemeFontScheme(A.FontScheme aFontScheme)
    {
        this.AFontScheme = aFontScheme;
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

    internal A.FontScheme AFontScheme { get; }

    private string GetBodyFontName()
    {
        return this.AFontScheme.MinorFont!.LatinFont!.Typeface!.Value!;
    }
    
    private void SetHeadingFontName(string fontName)
    {
        this.AFontScheme.MajorFont!.LatinFont!.Typeface!.Value = fontName;
    }

    private void SetBodyFontName(string fontName)
    {
        this.AFontScheme.MinorFont!.LatinFont!.Typeface!.Value = fontName;
    }
    
    private string GetHeadingFontName()
    {
        return this.AFontScheme.MajorFont!.LatinFont!.Typeface!.Value!;
    }
}
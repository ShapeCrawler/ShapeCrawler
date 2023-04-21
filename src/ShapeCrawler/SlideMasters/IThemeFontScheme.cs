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
    internal ThemeFontScheme(A.FontScheme aFontScheme)
    {
        this.AFontScheme = aFontScheme;
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

    internal A.FontScheme AFontScheme { get; }
    
    private string GetHeadLatinFont()
    {
        return this.AFontScheme.MajorFont!.LatinFont!.Typeface!.Value!;
    }
    
    private string GetHeadEastAsianFont()
    {
        return this.AFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value!;
    }
    
    private void SetHeadLatinFont(string fontName)
    {
        this.AFontScheme.MajorFont!.LatinFont!.Typeface!.Value = fontName;
    }
    
    private void SetHeadEastAsianFont(string fontName)
    {
        this.AFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value = fontName;
    }
    
    private string GetBodyLatinFont()
    {
        return this.AFontScheme.MinorFont!.LatinFont!.Typeface!.Value!;
    }
    
    private string GetBodyEastAsianFont()
    {
        return this.AFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value!;
    }
    
    private void SetBodyLatinFont(string fontName)
    {
        this.AFontScheme.MinorFont!.LatinFont!.Typeface!.Value = fontName;
    }
    
    private void SetBodyEastAsianFont(string fontName)
    {
        this.AFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value = fontName;
    }
}
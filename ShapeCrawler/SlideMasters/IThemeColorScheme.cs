using ShapeCrawler.Drawing;

namespace ShapeCrawler;

using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a color scheme.
/// </summary>
public interface IThemeColorScheme
{
    /// <summary>
    ///     Gets Dark 1 color in hexadecimal format.
    /// </summary>
    string Dark1 { get; }

    /// <summary>
    ///     Gets Light 1 color in hexadecimal format.
    /// </summary>
    string Light1 { get; }

    /// <summary>
    ///     Gets Dark 2 color in hexadecimal format.
    /// </summary>
    string Dark2 { get; }
    
    /// <summary>
    ///     Gets Light 2 color in hexadecimal format.
    /// </summary>
    string Light2 { get; }
    
    /// <summary>
    ///     Gets Accent 1 color in hexadecimal format.
    /// </summary>
    string Accent1 { get; }
    
    /// <summary>
    ///     Gets Accent 2 color in hexadecimal format.
    /// </summary>
    string Accent2 { get; }
    
    /// <summary>
    ///     Gets Accent 3 color in hexadecimal format.
    /// </summary>
    string Accent3 { get; }
    
    /// <summary>
    ///     Gets Accent 4 color in hexadecimal format.
    /// </summary>
    string Accent4 { get; }
    
    /// <summary>
    ///     Gets Accent 5 color in hexadecimal format.
    /// </summary>
    string Accent5 { get; }
    
    /// <summary>
    ///     Gets Accent 6 color in hexadecimal format.
    /// </summary>
    string Accent6 { get; }
    
    /// <summary>
    ///     Gets Hyperlink color in hexadecimal format.
    /// </summary>
    string Hyperlink { get; }
    
    /// <summary>
    ///     Gets Followed Hyperlink color in hexadecimal format.
    /// </summary>
    string FollowedHyperlink { get; }
}

internal class ThemeColorScheme : IThemeColorScheme
{
    private readonly A.ColorScheme aColorScheme;

    internal ThemeColorScheme(A.ColorScheme aColorScheme)
    {
        this.aColorScheme = aColorScheme;
    }

    public string Dark1 => this.GetColor(this.aColorScheme.Dark1Color!);
    
    public string Light1 => this.GetColor(this.aColorScheme.Light1Color!);
    
    public string Dark2 => this.GetColor(this.aColorScheme.Dark2Color!);

    public string Light2 => this.GetColor(this.aColorScheme.Light2Color!);

    public string Accent1 => this.GetColor(this.aColorScheme.Accent1Color!);

    public string Accent2 => this.GetColor(this.aColorScheme.Accent2Color!);

    public string Accent3 => this.GetColor(this.aColorScheme.Accent3Color!);

    public string Accent4 => this.GetColor(this.aColorScheme.Accent4Color!);
    
    public string Accent5 => this.GetColor(this.aColorScheme.Accent5Color!);
    
    public string Accent6 => this.GetColor(this.aColorScheme.Accent6Color!);
    
    public string Hyperlink => this.GetColor(this.aColorScheme.Hyperlink!);
    
    public string FollowedHyperlink => this.GetColor(this.aColorScheme.FollowedHyperlinkColor!);

    private string GetColor(A.Color2Type aColor2Type)
    {
        var color = HexParser.GetWithoutScheme(aColor2Type);
        return color!.Value.Item2;
    }
}
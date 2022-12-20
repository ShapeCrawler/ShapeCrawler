using System.Linq;
using ShapeCrawler.Drawing;

namespace ShapeCrawler;

using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a color scheme.
/// </summary>
public interface IThemeColorScheme
{
    /// <summary>
    ///     Gets or sets Dark 1 color in hexadecimal format.
    /// </summary>
    string Dark1 { get; set; }

    /// <summary>
    ///     Gets or sets Light 1 color in hexadecimal format.
    /// </summary>
    string Light1 { get; set; }

    /// <summary>
    ///     Gets or sets Dark 2 color in hexadecimal format.
    /// </summary>
    string Dark2 { get; set; }
    
    /// <summary>
    ///     Gets or sets Light 2 color in hexadecimal format.
    /// </summary>
    string Light2 { get; set; }
    
    /// <summary>
    ///     Gets or sets Accent 1 color in hexadecimal format.
    /// </summary>
    string Accent1 { get; set; }
    
    /// <summary>
    ///     Gets or sets Accent 2 color in hexadecimal format.
    /// </summary>
    string Accent2 { get; set; }
    
    /// <summary>
    ///     Gets or sets Accent 3 color in hexadecimal format.
    /// </summary>
    string Accent3 { get; set; }
    
    /// <summary>
    ///     Gets or sets Accent 4 color in hexadecimal format.
    /// </summary>
    string Accent4 { get; set; }
    
    /// <summary>
    ///     Gets or sets Accent 5 color in hexadecimal format.
    /// </summary>
    string Accent5 { get; set; }
    
    /// <summary>
    ///     Gets or sets Accent 6 color in hexadecimal format.
    /// </summary>
    string Accent6 { get; set; }
    
    /// <summary>
    ///     Gets or sets Hyperlink color in hexadecimal format.
    /// </summary>
    string Hyperlink { get; set; }
    
    /// <summary>
    ///     Gets or sets Followed Hyperlink color in hexadecimal format.
    /// </summary>
    string FollowedHyperlink { get; set; }
}

internal class ThemeColorScheme : IThemeColorScheme
{
    private readonly A.ColorScheme aColorScheme;

    internal ThemeColorScheme(A.ColorScheme aColorScheme)
    {
        this.aColorScheme = aColorScheme;
    }

    public string Dark1
    {
        get => this.GetColor(this.aColorScheme.Dark1Color!);
        set => this.SetColor("dk1", value);
    }

    public string Light1
    {
        get => this.GetColor(this.aColorScheme.Light1Color!);
        set => this.SetColor("lt1", value);
    }

    public string Dark2
    {
        get => this.GetColor(this.aColorScheme.Dark2Color!);
        set => this.SetColor("dk2", value);
    }

    public string Light2
    {
        get => this.GetColor(this.aColorScheme.Light2Color!);
        set => this.SetColor("lt2", value);
    }

    public string Accent1
    {
        get => this.GetColor(this.aColorScheme.Accent1Color!);
        set => this.SetColor("accent1", value);
    }

    public string Accent2
    {
        get => this.GetColor(this.aColorScheme.Accent2Color!);
        set => this.SetColor("accent2", value);
    }

    public string Accent3
    {
        get => this.GetColor(this.aColorScheme.Accent3Color!);
        set => this.SetColor("accent3", value);
    }

    public string Accent4
    {
        get => this.GetColor(this.aColorScheme.Accent4Color!);
        set => this.SetColor("accent4", value);
    }

    public string Accent5
    {
        get => this.GetColor(this.aColorScheme.Accent5Color!);
        set => this.SetColor("accent5", value);
    }

    public string Accent6
    {
        get => this.GetColor(this.aColorScheme.Accent6Color!);
        set => this.SetColor("accent6", value);
    }

    public string Hyperlink
    {
        get => this.GetColor(this.aColorScheme.Hyperlink!);
        set => this.SetColor("hlink", value);
    }

    public string FollowedHyperlink
    {
        get => this.GetColor(this.aColorScheme.FollowedHyperlinkColor!);
        set => this.SetColor("folHlink", value);
    }

    private string GetColor(A.Color2Type aColor2Type)
    {
        var color = HexParser.GetWithoutScheme(aColor2Type);
        return color!.Value.Item2;
    }
    
    private void SetColor(string name, string hex)
    {
        var color = this.aColorScheme.Elements().First(x => x.LocalName == name);
        foreach (var child in color)
        {
            child.Remove();
        }
        
        var aSrgbClr = new A.RgbColorModelHex { Val = hex };
        color.Append(aSrgbClr);
    }
}
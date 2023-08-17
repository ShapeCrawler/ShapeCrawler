// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a theme.
/// </summary>
public interface ITheme
{
    /// <summary>
    ///     Gets font scheme.
    /// </summary>
    IThemeFontScheme FontScheme { get; }

    /// <summary>
    ///     Gets color scheme.
    /// </summary>
    IThemeColorScheme ColorScheme { get; }
}

internal sealed class SCTheme : ITheme
{
    private readonly A.Theme aTheme;

    internal SCTheme(SlideMaster parentMaster, A.Theme aTheme)
    {
        this.aTheme = aTheme;
    }

    public IThemeFontScheme FontScheme => this.GetFontSetting();

    public IThemeColorScheme ColorScheme => this.GetColorScheme();

    private IThemeFontScheme GetFontSetting()
    {
        return new ThemeFontScheme(this.aTheme.ThemeElements!.FontScheme!);
    }
    
    private IThemeColorScheme GetColorScheme()
    {
        return new ThemeColorScheme(this.aTheme.ThemeElements!.ColorScheme!);
    }
}
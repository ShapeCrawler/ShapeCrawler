using System;
using ShapeCrawler.Drawing;

namespace ShapeCrawler;

using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a color scheme.
/// </summary>
public interface IThemeColorScheme
{
    /// <summary>
    ///     Gets Dark1 color in hexadecimal format.
    /// </summary>
    string Dark1 { get; }
}

internal class ThemeColorScheme : IThemeColorScheme
{
    private readonly A.ColorScheme aColorScheme;

    internal ThemeColorScheme(A.ColorScheme aColorScheme)
    {
        this.aColorScheme = aColorScheme;
    }

    public string Dark1 => this.GetDark1();

    private string GetDark1()
    {
        var color = HexParser.GetWithoutScheme(this.aColorScheme.Dark1Color!);
        
        return color!.Value.Item2; // TODO: remove "!"
    }
}
using System;
using System.Linq;
using System.Reflection;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Colors;

internal static class ColorTranslator
{
    private static readonly FieldInfo[] FieldInfoList = typeof(SKColors).GetFields(BindingFlags.Static | BindingFlags.Public);

    internal static string? NameFromSchemeColor(A.SchemeColor? schemeColor)
    {
        return schemeColor?.Val?.InnerText switch
        {
            "bg1" => "Background 1",
            "tx1" => "Text 1",
            "bg2" => "Background 2",
            "tx2" => "Text 2",
            "dk1" => "Dark 1",
            "lt1" => "Light 1",
            "dk2" => "Dark 2",
            "lt2" => "Light 2",
            "accent1" => "Accent 1",
            "accent2" => "Accent 2",
            "accent3" => "Accent 3",
            "accent4" => "Accent 4",
            "accent5" => "Accent 5",
            "accent6" => "Accent 6",
            "hlink" => "Hyperlink",
            "folHlink" => "Followed Hyperlink",
            "phClr" => "Placeholder",
            _ => null
        };
    }

    internal static string HexFromName(string colorName)
    {
        if (colorName.Equals("white", StringComparison.CurrentCultureIgnoreCase))
        {
            return "FFFFFF";
        }

        var fieldInfo = FieldInfoList.First(fieldInfo => string.Equals(fieldInfo.Name, colorName, StringComparison.CurrentCultureIgnoreCase));
        var color = (SKColor)fieldInfo.GetValue(null)!;

        return color.ToString();
    }
}

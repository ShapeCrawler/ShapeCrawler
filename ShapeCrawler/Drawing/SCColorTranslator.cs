using System;
using System.Linq;
using System.Reflection;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

internal static class SCColorTranslator
{
    private static readonly FieldInfo[] fieldInfoList;

    static SCColorTranslator()
    {
        fieldInfoList = typeof(SKColors).GetFields(BindingFlags.Static | BindingFlags.Public);
    }

    internal static string HexFromName(string coloName)
    {
        if (coloName.ToLower() == "white")
        {
            return "FFFFFF";
        }

        var fieldInfo = fieldInfoList.First(fieldInfo => string.Equals(fieldInfo.Name, coloName, StringComparison.CurrentCultureIgnoreCase));
        var color = (SKColor)fieldInfo.GetValue(null) !;

        return color.ToString();
    }
}
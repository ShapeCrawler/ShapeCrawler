using System;
using System.Collections.Generic;

namespace ShapeCrawler.Tables;

// TODO : verify the values 
internal static class CommonTableStyles
{
    private static readonly Dictionary<string, ITableStyle> Styles = new()
    {
        // https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh273476(v=office.14)?redirectedfrom=MSDN
        { "No Style, No Grid", TableStyle.NoStyleNoGrid },
        { "No Style, Table Grid", TableStyle.NoStyleTableGrid },
        { "Themed Style 1 - Accent 1", TableStyle.ThemedStyle1Accent1 },
        { "Themed Style 1 - Accent 2", TableStyle.ThemedStyle1Accent2 },
        { "Themed Style 1 - Accent 3", TableStyle.ThemedStyle1Accent3 },
        { "Themed Style 1 - Accent 4", TableStyle.ThemedStyle1Accent4 },
        { "Themed Style 1 - Accent 5", TableStyle.ThemedStyle1Accent5 },
        { "Themed Style 1 - Accent 6", TableStyle.ThemedStyle1Accent6 },
        { "Themed Style 2 - Accent 1", TableStyle.ThemedStyle2Accent1 },
        { "Themed Style 2 - Accent 2", TableStyle.ThemedStyle2Accent2 },
        { "Themed Style 2 - Accent 3", TableStyle.ThemedStyle2Accent3 },
        { "Themed Style 2 - Accent 4", TableStyle.ThemedStyle2Accent4 },
        { "Themed Style 2 - Accent 5", TableStyle.ThemedStyle2Accent5 },
        { "Themed Style 2 - Accent 6", TableStyle.ThemedStyle2Accent6 },
        { "Light Style 1", TableStyle.LightStyle1 },
        { "Light Style 1 - Accent 1", TableStyle.LightStyle1Accent1 },
        { "Light Style 1 - Accent 2", TableStyle.LightStyle1Accent2 },
        { "Light Style 1 - Accent 3", TableStyle.LightStyle1Accent3 },
        { "Light Style 1 - Accent 4", TableStyle.LightStyle1Accent4 },
        { "Light Style 1 - Accent 5", TableStyle.LightStyle1Accent5 },
        { "Light Style 1 - Accent 6", TableStyle.LightStyle1Accent6 },
        { "Light Style 2", TableStyle.LightStyle2 },
        { "Light Style 2 - Accent 1", TableStyle.LightStyle2Accent1 },
        { "Light Style 2 - Accent 2", TableStyle.LightStyle2Accent2 },
        { "Light Style 2 - Accent 3", TableStyle.LightStyle2Accent3 },
        { "Light Style 2 - Accent 4", TableStyle.LightStyle2Accent4 },
        { "Light Style 2 - Accent 5", TableStyle.LightStyle2Accent5 },
        { "Light Style 2 - Accent 6", TableStyle.LightStyle2Accent6 },
        { "Light Style 3", TableStyle.LightStyle3 },
        { "Light Style 3 - Accent 1", TableStyle.LightStyle3Accent1 },
        { "Light Style 3 - Accent 2", TableStyle.LightStyle3Accent2 },
        { "Light Style 3 - Accent 3", TableStyle.LightStyle3Accent3 },
        { "Light Style 3 - Accent 4", TableStyle.LightStyle3Accent4 },
        { "Light Style 3 - Accent 5", TableStyle.LightStyle3Accent5 },
        { "Light Style 3 - Accent 6", TableStyle.LightStyle3Accent6 },
        { "Medium Style 1", TableStyle.MediumStyle1 },
        { "Medium Style 1 - Accent 1", TableStyle.MediumStyle1Accent1 },
        { "Medium Style 1 - Accent 2", TableStyle.MediumStyle1Accent2 },
        { "Medium Style 1 - Accent 3", TableStyle.MediumStyle1Accent3 },
        { "Medium Style 1 - Accent 4", TableStyle.MediumStyle1Accent4 },
        { "Medium Style 1 - Accent 5", TableStyle.MediumStyle1Accent5 },
        { "Medium Style 1 - Accent 6", TableStyle.MediumStyle1Accent6 },
        { "Medium Style 2", TableStyle.MediumStyle2 },
        { "Medium Style 2 - Accent 1", TableStyle.MediumStyle2Accent1 },
        { "Medium Style 2 - Accent 2", TableStyle.MediumStyle2Accent2 },
        { "Medium Style 2 - Accent 3", TableStyle.MediumStyle2Accent3 },
        { "Medium Style 2 - Accent 4", TableStyle.MediumStyle2Accent4 },
        { "Medium Style 2 - Accent 5", TableStyle.MediumStyle2Accent5 },
        { "Medium Style 2 - Accent 6", TableStyle.MediumStyle2Accent6 },
        { "Medium Style 3", TableStyle.MediumStyle3 },
        { "Medium Style 3 - Accent 1", TableStyle.MediumStyle3Accent1 },
        { "Medium Style 3 - Accent 2", TableStyle.MediumStyle3Accent2 },
        { "Medium Style 3 - Accent 3", TableStyle.MediumStyle3Accent3 },
        { "Medium Style 3 - Accent 4", TableStyle.MediumStyle3Accent4 },
        { "Medium Style 3 - Accent 5", TableStyle.MediumStyle3Accent5 },
        { "Medium Style 3 - Accent 6", TableStyle.MediumStyle3Accent6 },
        { "Medium Style 4", TableStyle.MediumStyle4 },
        { "Medium Style 4 - Accent 1", TableStyle.MediumStyle4Accent1 },
        { "Medium Style 4 - Accent 2", TableStyle.MediumStyle4Accent2 },
        { "Medium Style 4 - Accent 3", TableStyle.MediumStyle4Accent3 },
        { "Medium Style 4 - Accent 4", TableStyle.MediumStyle4Accent4 },
        { "Medium Style 4 - Accent 5", TableStyle.MediumStyle4Accent5 },
        { "Medium Style 4 - Accent 6", TableStyle.MediumStyle4Accent6 },
        { "Dark Style 1", TableStyle.DarkStyle1 },
        { "Dark Style 1 - Accent 1", TableStyle.DarkStyle1Accent1 },
        { "Dark Style 1 - Accent 2", TableStyle.DarkStyle1Accent2 },
        { "Dark Style 1 - Accent 3", TableStyle.DarkStyle1Accent3 },
        { "Dark Style 1 - Accent 5", TableStyle.DarkStyle1Accent5 },
        { "Dark Style 1 - Accent 4", TableStyle.DarkStyle1Accent4 },
        { "Dark Style 1 - Accent 6", TableStyle.DarkStyle1Accent6 },
        { "Dark Style 2", TableStyle.DarkStyle2 },
        { "Dark Style 2 - Accent 1, Accent 2", TableStyle.DarkStyle2Accent1Accent2 },
        { "Dark Style 2 - Accent 3, Accent 4", TableStyle.DarkStyle2Accent3Accent4 }, 
        { "Dark Style 2 - Accent 5, Accent 6", TableStyle.DarkStyle2Accent5Accent6 }
    };
    
    /// <summary>
    ///     Get the style using its GUID.
    /// </summary>
    public static ITableStyle? GetTableStyleByGuid(string guid)
    {
        // Search through the dictionary for the matching GUID
        foreach (var value in Styles)
        {
            if (value.Value.Guid.Equals(guid, StringComparison.OrdinalIgnoreCase))
            {
                return value.Value;
            }
        }

        return null;
    }
}
using System;
using System.Collections.Generic;

namespace ShapeCrawler.Tables;

// TODO : verify the values 

/// <summary>
///     Represents the common table style.
/// </summary>
public enum TableStyleEnum
{
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
// enum is self documented
#pragma warning disable SA1602 // Enumeration items should be documented
    NoStyleNoGrid,
    NoStyleTableGrid,

    ThemedStyle1Accent1,
    ThemedStyle1Accent2,
    ThemedStyle1Accent3,
    ThemedStyle1Accent4,
    ThemedStyle1Accent5,
    ThemedStyle1Accent6,

    ThemedStyle2Accent1,
    ThemedStyle2Accent2,
    ThemedStyle2Accent3,
    ThemedStyle2Accent4,
    ThemedStyle2Accent5,
    ThemedStyle2Accent6,

    LightStyle1,
    LightStyle1Accent1,
    LightStyle1Accent2,
    LightStyle1Accent3,
    LightStyle1Accent4,
    LightStyle1Accent5,
    LightStyle1Accent6,

    LightStyle2,
    LightStyle2Accent1,
    LightStyle2Accent2,
    LightStyle2Accent3,
    LightStyle2Accent4,
    LightStyle2Accent5,
    LightStyle2Accent6,

    LightStyle3,
    LightStyle3Accent1,
    LightStyle3Accent2,
    LightStyle3Accent3,
    LightStyle3Accent4,
    LightStyle3Accent5,
    LightStyle3Accent6,

    MediumStyle1,
    MediumStyle1Accent1,
    MediumStyle1Accent2,
    MediumStyle1Accent3,
    MediumStyle1Accent4,
    MediumStyle1Accent5,
    MediumStyle1Accent6,

    MediumStyle2,
    MediumStyle2Accent1,
    MediumStyle2Accent2,
    MediumStyle2Accent3,
    MediumStyle2Accent4,
    MediumStyle2Accent5,
    MediumStyle2Accent6,

    MediumStyle3,
    MediumStyle3Accent1,
    MediumStyle3Accent2,
    MediumStyle3Accent3,
    MediumStyle3Accent4,
    MediumStyle3Accent5,
    MediumStyle3Accent6,

    MediumStyle4,
    MediumStyle4Accent1,
    MediumStyle4Accent2,
    MediumStyle4Accent3,
    MediumStyle4Accent4,
    MediumStyle4Accent5,
    MediumStyle4Accent6,

    DarkStyle1,
    DarkStyle1Accent1,
    DarkStyle1Accent2,
    DarkStyle1Accent3,
    DarkStyle1Accent4,
    DarkStyle1Accent5,
    DarkStyle1Accent6,

    DarkStyle2,
    DarkStyle2Accent1Accent2,
    DarkStyle2Accent3Accent4,
    DarkStyle2Accent5Accent6
#pragma warning restore SA1602 // Enumeration items should be documented

#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
}

internal static class CommonTableStyles
{
    public static readonly Dictionary<TableStyleEnum, TableStyle> Styles = new Dictionary<TableStyleEnum, TableStyle>
    {
        { TableStyleEnum.NoStyleNoGrid, new TableStyle("No Style, No Grid", "{2D5ABB26-0587-4C30-8999-92F81FD0307C}") },
        { TableStyleEnum.NoStyleTableGrid, new TableStyle("No Style, Table Grid", "{5940675A-B579-460E-94D1-54222C63F5DA}") },
        { TableStyleEnum.ThemedStyle1Accent1, new TableStyle("Themed Style 1 - Accent 1", "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}") },
        { TableStyleEnum.ThemedStyle1Accent2, new TableStyle("Themed Style 1 - Accent 2", "{284E427A-3D55-4303-BF80-6455036E1DE7}") },
        { TableStyleEnum.ThemedStyle1Accent3, new TableStyle("Themed Style 1 - Accent 3", "{69C7853C-536D-4A76-A0AE-DD22124D55A5}") },
        { TableStyleEnum.ThemedStyle1Accent4, new TableStyle("Themed Style 1 - Accent 4", "{775DCB02-9BB8-47FD-8907-85C794F793BA}") },
        { TableStyleEnum.ThemedStyle1Accent5, new TableStyle("Themed Style 1 - Accent 5", "{35758FB7-9AC5-4552-8A53-C91805E547FA}") },
        { TableStyleEnum.ThemedStyle1Accent6, new TableStyle("Themed Style 1 - Accent 6", "{08FB837D-C827-4EFA-A057-4D05807E0F7C}") },
        { TableStyleEnum.ThemedStyle2Accent1, new TableStyle("Themed Style 2 - Accent 1", "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}") },
        { TableStyleEnum.ThemedStyle2Accent2, new TableStyle("Themed Style 2 - Accent 2", "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}") },
        { TableStyleEnum.ThemedStyle2Accent3, new TableStyle("Themed Style 2 - Accent 3", "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}") },
        { TableStyleEnum.ThemedStyle2Accent4, new TableStyle("Themed Style 2 - Accent 4", "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}") },
        { TableStyleEnum.ThemedStyle2Accent5, new TableStyle("Themed Style 2 - Accent 5", "{327F97BB-C833-4FB7-BDE5-3F7075034690}") },
        { TableStyleEnum.ThemedStyle2Accent6, new TableStyle("Themed Style 2 - Accent 6", "{638B1855-1B75-4FBE-930C-398BA8C253C6}") },
        { TableStyleEnum.LightStyle1, new TableStyle("Light Style 1", "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}") },
        { TableStyleEnum.LightStyle1Accent1, new TableStyle("Light Style 1 - Accent 1", "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}") },
        { TableStyleEnum.LightStyle1Accent2, new TableStyle("Light Style 1 - Accent 2", "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}") },
        { TableStyleEnum.LightStyle1Accent3, new TableStyle("Light Style 1 - Accent 3", "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}") },
        { TableStyleEnum.LightStyle1Accent4, new TableStyle("Light Style 1 - Accent 4", "{D27102A9-8310-4765-A935-A1911B00CA55}") },
        { TableStyleEnum.LightStyle1Accent5, new TableStyle("Light Style 1 - Accent 5", "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}") },
        { TableStyleEnum.LightStyle1Accent6, new TableStyle("Light Style 1 - Accent 6", "{68D230F3-CF80-4859-8CE7-A43EE81993B5}") },
        { TableStyleEnum.LightStyle2, new TableStyle("Light Style 2", "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}") },
        { TableStyleEnum.LightStyle2Accent1, new TableStyle("Light Style 2 - Accent 1", "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}") },
        { TableStyleEnum.LightStyle2Accent2, new TableStyle("Light Style 2 - Accent 2", "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}") },
        { TableStyleEnum.LightStyle2Accent3, new TableStyle("Light Style 2 - Accent 3", "{F2DE63D5-997A-4646-A377-4702673A728D}") },
        { TableStyleEnum.LightStyle2Accent4, new TableStyle("Light Style 2 - Accent 4", "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}") },
        { TableStyleEnum.LightStyle2Accent5, new TableStyle("Light Style 2 - Accent 5", "{5A111915-BE36-4E01-A7E5-04B1672EAD32}") },
        { TableStyleEnum.LightStyle2Accent6, new TableStyle("Light Style 2 - Accent 6", "{912C8C85-51F0-491E-9774-3900AFEF0FD7}") },
        { TableStyleEnum.LightStyle3, new TableStyle("Light Style 3", "{616DA210-FB5B-4158-B5E0-FEB733F419BA}") },
        { TableStyleEnum.LightStyle3Accent1, new TableStyle("Light Style 3 - Accent 1", "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}") },
        { TableStyleEnum.LightStyle3Accent2, new TableStyle("Light Style 3 - Accent 2", "{5DA37D80-6434-44D0-A028-1B22A696006F}") },
        { TableStyleEnum.LightStyle3Accent3, new TableStyle("Light Style 3 - Accent 3", "{8799B23B-EC83-4686-B30A-512413B5E67A}") },
        { TableStyleEnum.LightStyle3Accent4, new TableStyle("Light Style 3 - Accent 4", "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}") },
        { TableStyleEnum.LightStyle3Accent5, new TableStyle("Light Style 3 - Accent 5", "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}") },
        { TableStyleEnum.LightStyle3Accent6, new TableStyle("Light Style 3 - Accent 6", "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}") },
        { TableStyleEnum.MediumStyle1, new TableStyle("Medium Style 1", "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}") },
        { TableStyleEnum.MediumStyle1Accent1, new TableStyle("Medium Style 1 - Accent 1", "{B301B821-A1FF-4177-AEE7-76D212191A09}") },
        { TableStyleEnum.MediumStyle1Accent2, new TableStyle("Medium Style 1 - Accent 2", "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}") },
        { TableStyleEnum.MediumStyle1Accent3, new TableStyle("Medium Style 1 - Accent 3", "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}") },
        { TableStyleEnum.MediumStyle1Accent4, new TableStyle("Medium Style 1 - Accent 4", "{1E171933-4619-4E11-9A3F-F7608DF75F80}") },
        { TableStyleEnum.MediumStyle1Accent5, new TableStyle("Medium Style 1 - Accent 5", "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}") },
        { TableStyleEnum.MediumStyle1Accent6, new TableStyle("Medium Style 1 - Accent 6", "{10A1B5D5-9B99-4C35-A422-299274C87663}") },
        { TableStyleEnum.MediumStyle2, new TableStyle("Medium Style 2", "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}") },
        { TableStyleEnum.MediumStyle2Accent1, new TableStyle("Medium Style 2 - Accent 1", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}") },
        { TableStyleEnum.MediumStyle2Accent2, new TableStyle("Medium Style 2 - Accent 2", "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}") },
        { TableStyleEnum.MediumStyle2Accent3, new TableStyle("Medium Style 2 - Accent 3", "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}") },
        { TableStyleEnum.MediumStyle2Accent4, new TableStyle("Medium Style 2 - Accent 4", "{00A15C55-8517-42AA-B614-E9B94910E393}") },
        { TableStyleEnum.MediumStyle2Accent5, new TableStyle("Medium Style 2 - Accent 5", "{7DF18680-E054-41AD-8BC1-D1AEF772440D}") },
        { TableStyleEnum.MediumStyle2Accent6, new TableStyle("Medium Style 2 - Accent 6", "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}") },
        { TableStyleEnum.MediumStyle3, new TableStyle("Medium Style 3", "{8EC20E35-A176-4012-BC5E-935CFFF8708E}") },
        { TableStyleEnum.MediumStyle3Accent1, new TableStyle("Medium Style 3 - Accent 1", "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}") },
        { TableStyleEnum.MediumStyle3Accent2, new TableStyle("Medium Style 3 - Accent 2", "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}") },
        { TableStyleEnum.MediumStyle3Accent3, new TableStyle("Medium Style 3 - Accent 3", "{EB344D84-9AFB-497E-A393-DC336BA19D2E}") },
        { TableStyleEnum.MediumStyle3Accent4, new TableStyle("Medium Style 3 - Accent 4", "{EB9631B5-78F2-41C9-869B-9F39066F8104}") },
        { TableStyleEnum.MediumStyle3Accent5, new TableStyle("Medium Style 3 - Accent 5", "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}") },
        { TableStyleEnum.MediumStyle3Accent6, new TableStyle("Medium Style 3 - Accent 6", "{2A488322-F2BA-4B5B-9748-0D474271808F}") },
        { TableStyleEnum.MediumStyle4, new TableStyle("Medium Style 4", "{D7AC3CCA-C797-4891-BE02-D94E43425B78}") },
        { TableStyleEnum.MediumStyle4Accent1, new TableStyle("Medium Style 4 - Accent 1", "{69CF1AB2-1976-4502-BF36-3FF5EA218861}") },
        { TableStyleEnum.MediumStyle4Accent2, new TableStyle("Medium Style 4 - Accent 2", "{8A107856-5554-42FB-B03E-39F5DBC370BA}") },
        { TableStyleEnum.MediumStyle4Accent3, new TableStyle("Medium Style 4 - Accent 3", "{0505E3EF-67EA-436B-97B2-0124C06EBD24}") },
        { TableStyleEnum.MediumStyle4Accent4, new TableStyle("Medium Style 4 - Accent 4", "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}") },
        { TableStyleEnum.MediumStyle4Accent5, new TableStyle("Medium Style 4 - Accent 5", "{22838BEF-8BB2-4498-84A7-C5851F593DF1}") },
        { TableStyleEnum.MediumStyle4Accent6, new TableStyle("Medium Style 4 - Accent 6", "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}") },
        { TableStyleEnum.DarkStyle1, new TableStyle("Dark Style 1", "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}") },
        { TableStyleEnum.DarkStyle1Accent1, new TableStyle("Dark Style 1 - Accent 1", "{125E5076-3810-47DD-B79F-674D7AD40C01}") },
        { TableStyleEnum.DarkStyle1Accent2, new TableStyle("Dark Style 1 - Accent 2", "{37CE84F3-28C3-443E-9E96-99CF82512B78}") },
        { TableStyleEnum.DarkStyle1Accent3, new TableStyle("Dark Style 1 - Accent 3", "{D03447BB-5D67-496B-8E87-E561075AD55C}") },
        { TableStyleEnum.DarkStyle1Accent4, new TableStyle("Dark Style 1 - Accent 4", "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}") },
        { TableStyleEnum.DarkStyle1Accent5, new TableStyle("Dark Style 1 - Accent 5", "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}") },
        { TableStyleEnum.DarkStyle1Accent6, new TableStyle("Dark Style 1 - Accent 6", "{AF606853-7671-496A-8E4F-DF71F8EC918B}") },
        { TableStyleEnum.DarkStyle2, new TableStyle("Dark Style 2", "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}") },
        { TableStyleEnum.DarkStyle2Accent1Accent2, new TableStyle("Dark Style 2 - Accent 1, Accent 2", "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}") },
        { TableStyleEnum.DarkStyle2Accent3Accent4, new TableStyle("Dark Style 2 - Accent 3, Accent 4", "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}") },
        { TableStyleEnum.DarkStyle2Accent5Accent6, new TableStyle("Dark Style 2 - Accent 5, Accent 6", "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}") }
    };



    /// <summary>
    ///     Gets the style using enum.
    /// </summary>
    public static ITableStyle GetTableStyleByEnum(TableStyleEnum enum_value)
    {
        return CommonTableStyles.Styles[enum_value];
    }

    /// <summary>
    ///     Get the style using its name.
    /// </summary>
    public static ITableStyle? GetTableStyleByName(string name)
    {
        // Search through the dictionary for the matching name
        foreach (var value in CommonTableStyles.Styles)
        {
            if (value.Value.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return value.Value;
            }
        }

        // If no matching table style is found, return null
        return null;
    }

    /// <summary>
    ///     Get the style using its GUID.
    /// </summary>
    public static ITableStyle? GetTableStyleByGUID(string guid)
    {
        // Search through the dictionary for the matching GUID
        foreach (var value in CommonTableStyles.Styles)
        {
            if (value.Value.GUID.Equals(guid, StringComparison.OrdinalIgnoreCase))
            {
                return value.Value;
            }
        }

        // If no matching table style is found, return null
        return null;
    }

    /// <summary>
    ///     Get the style either by name or GUID.
    /// </summary>
    public static ITableStyle? SearchTableStyle(string search)
    {
        var res = GetTableStyleByName(search);

        if (res == null)
        {
            res = GetTableStyleByGUID(search);
        }

        return res;
    }
}


/*

//list found here : https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh273476(v=office.14)?redirectedfrom=MSDN

' No Style, No Grid: {2D5ABB26-0587-4C30-8999-92F81FD0307C}
' Themed Style 1 - Accent 1: {3C2FFA5D-87B4-456A-9821-1D502468CF0F}
' Themed Style 1 - Accent 2: {284E427A-3D55-4303-BF80-6455036E1DE7}
' Themed Style 1 - Accent 3: {69C7853C-536D-4A76-A0AE-DD22124D55A5}
' Themed Style 1 - Accent 4: {775DCB02-9BB8-47FD-8907-85C794F793BA}
' Themed Style 1 - Accent 5: {35758FB7-9AC5-4552-8A53-C91805E547FA}
' Themed Style 1 - Accent 6: {08FB837D-C827-4EFA-A057-4D05807E0F7C}
' No Style, Table Grid: {5940675A-B579-460E-94D1-54222C63F5DA}
' Themed Style 2 - Accent 1: {D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}
' Themed Style 2 - Accent 2: {18603FDC-E32A-4AB5-989C-0864C3EAD2B8}
' Themed Style 2 - Accent 3: {306799F8-075E-4A3A-A7F6-7FBC6576F1A4}
' Themed Style 2 - Accent 4: {E269D01E-BC32-4049-B463-5C60D7B0CCD2}
' Themed Style 2 - Accent 5: {327F97BB-C833-4FB7-BDE5-3F7075034690}
' Themed Style 2 - Accent 6: {638B1855-1B75-4FBE-930C-398BA8C253C6}
' Light Style 1: {9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}
' Light Style 1 - Accent 1: {3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}
' Light Style 1 - Accent 2: {0E3FDE45-AF77-4B5C-9715-49D594BDF05E}
' Light Style 1 - Accent 3: {C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}
' Light Style 1 - Accent 4: {D27102A9-8310-4765-A935-A1911B00CA55}
' Light Style 1 - Accent 5: {5FD0F851-EC5A-4D38-B0AD-8093EC10F338}
' Light Style 1 - Accent 6: {68D230F3-CF80-4859-8CE7-A43EE81993B5}
' Light Style 2: {7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}
' Light Style 2 - Accent 1: {69012ECD-51FC-41F1-AA8D-1B2483CD663E}
' Light Style 2 - Accent 2: {72833802-FEF1-4C79-8D5D-14CF1EAF98D9}
' Light Style 2 - Accent 3: {F2DE63D5-997A-4646-A377-4702673A728D}
' Light Style 2 - Accent 4: {17292A2E-F333-43FB-9621-5CBBE7FDCDCB}
' Light Style 2 - Accent 5: {5A111915-BE36-4E01-A7E5-04B1672EAD32}
' Light Style 2 - Accent 6: {912C8C85-51F0-491E-9774-3900AFEF0FD7}
' Light Style 3: {616DA210-FB5B-4158-B5E0-FEB733F419BA}
' Light Style 3 - Accent 1: {BC89EF96-8CEA-46FF-86C4-4CE0E7609802}
' Light Style 3 - Accent 2: {5DA37D80-6434-44D0-A028-1B22A696006F}
' Light Style 3 - Accent 3: {8799B23B-EC83-4686-B30A-512413B5E67A}
' Light Style 3 - Accent 4: {ED083AE6-46FA-4A59-8FB0-9F97EB10719F}
' Light Style 3 - Accent 5: {BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}
' Light Style 3 - Accent 6: {E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}
' Medium Style 1: {793D81CF-94F2-401A-BA57-92F5A7B2D0C5}
' Medium Style 1 - Accent 1: {B301B821-A1FF-4177-AEE7-76D212191A09}
' Medium Style 1 - Accent 2: {9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}
' Medium Style 1 - Accent 3: {1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}
' Medium Style 1 - Accent 4: {1E171933-4619-4E11-9A3F-F7608DF75F80}
' Medium Style 1 - Accent 5: {FABFCF23-3B69-468F-B69F-88F6DE6A72F2}
' Medium Style 1 - Accent 6: {10A1B5D5-9B99-4C35-A422-299274C87663}
' Medium Style 2: {073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}
' Medium Style 2 - Accent 1: {5C22544A-7EE6-4342-B048-85BDC9FD1C3A}
' Medium Style 2 - Accent 2: {21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}
' Medium Style 2 - Accent 3: {F5AB1C69-6EDB-4FF4-983F-18BD219EF322}
' Medium Style 2 - Accent 4: {00A15C55-8517-42AA-B614-E9B94910E393}
' Medium Style 2 - Accent 5: {7DF18680-E054-41AD-8BC1-D1AEF772440D}
' Medium Style 2 - Accent 6: {93296810-A885-4BE3-A3E7-6D5BEEA58F35}
' Medium Style 3: {8EC20E35-A176-4012-BC5E-935CFFF8708E}
' Medium Style 3 - Accent 1: {6E25E649-3F16-4E02-A733-19D2CDBF48F0}
' Medium Style 3 - Accent 2: {85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}
' Medium Style 3 - Accent 3: {EB344D84-9AFB-497E-A393-DC336BA19D2E}
' Medium Style 3 - Accent 4: {EB9631B5-78F2-41C9-869B-9F39066F8104}
' Medium Style 3 - Accent 5: {74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}
' Medium Style 3 - Accent 6: {2A488322-F2BA-4B5B-9748-0D474271808F}
' Medium Style 4: {D7AC3CCA-C797-4891-BE02-D94E43425B78}
' Medium Style 4 - Accent 1: {69CF1AB2-1976-4502-BF36-3FF5EA218861}
' Medium Style 4 - Accent 2: {8A107856-5554-42FB-B03E-39F5DBC370BA}
' Medium Style 4 - Accent 3: {0505E3EF-67EA-436B-97B2-0124C06EBD24}
' Medium Style 4 - Accent 4: {C4B1156A-380E-4F78-BDF5-A606A8083BF9}
' Medium Style 4 - Accent 5: {22838BEF-8BB2-4498-84A7-C5851F593DF1}
' Medium Style 4 - Accent 6: {16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}
' Dark Style 1: {E8034E78-7F5D-4C2E-B375-FC64B27BC917}
' Dark Style 1 - Accent 1: {125E5076-3810-47DD-B79F-674D7AD40C01}
' Dark Style 1 - Accent 2: {37CE84F3-28C3-443E-9E96-99CF82512B78}
' Dark Style 1 - Accent 3: {D03447BB-5D67-496B-8E87-E561075AD55C}
' Dark Style 1 - Accent 4: {E929F9F4-4A8F-4326-A1B4-22849713DDAB}
' Dark Style 1 - Accent 5:{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}
' Dark Style 1 - Accent 6: {AF606853-7671-496A-8E4F-DF71F8EC918B}
' Dark Style 2: {5202B0CA-FC54-4496-8BCA-5EF66A818D29}
' Dark Style 2 - Accent 1/Accent 2: {0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}
' Dark Style 2 - Accent 3/Accent 4: {91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}
' Dark Style 2 - Accent 5/Accent 6: {46F890A9-2807-4EBB-B81D-B2AA78EC7F39}

 */

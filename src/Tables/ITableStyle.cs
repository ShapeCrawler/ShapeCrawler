using System.Collections.Generic;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a table style of a table.
/// </summary>
public interface ITableStyle
{
    /// <summary>
    ///     Gets the name of the style.
    /// </summary> 
    public string Name { get; }
}

internal class TableStyle(string name): ITableStyle
{
    public string Name { get; } = name;

    public string Guid { get; private set; } = string.Empty;
    
    internal static TableStyle NoStyleNoGrid => new("No Style, No Grid") { Guid = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}" };

    internal static TableStyle NoStyleTableGrid => new("No Style, Table Grid") { Guid = "{5940675A-B579-460E-94D1-54222C63F5DA}" };

    internal static TableStyle ThemedStyle1Accent1 => new("Themed Style 1 - Accent 1") { Guid = "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}" };

    internal static TableStyle ThemedStyle1Accent2 => new("Themed Style 1 - Accent 2") { Guid = "{284E427A-3D55-4303-BF80-6455036E1DE7}" };

    internal static TableStyle ThemedStyle1Accent3 => new("Themed Style 1 - Accent 3") { Guid = "{69C7853C-536D-4A76-A0AE-DD22124D55A5}" };

    internal static TableStyle ThemedStyle1Accent4 => new("Themed Style 1 - Accent 4") { Guid = "{775DCB02-9BB8-47FD-8907-85C794F793BA}" };

    internal static TableStyle ThemedStyle1Accent5 => new("Themed Style 1 - Accent 5") { Guid = "{35758FB7-9AC5-4552-8A53-C91805E547FA}" };

    internal static TableStyle ThemedStyle1Accent6 => new("Themed Style 1 - Accent 6") { Guid = "{08FB837D-C827-4EFA-A057-4D05807E0F7C}" };

    internal static TableStyle ThemedStyle2Accent1 => new("Themed Style 2 - Accent 1") { Guid = "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}" };

    internal static TableStyle ThemedStyle2Accent2 => new("Themed Style 2 - Accent 2") { Guid = "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}" };

    internal static TableStyle ThemedStyle2Accent3 => new("Themed Style 2 - Accent 3") { Guid = "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}" };

    internal static TableStyle ThemedStyle2Accent4 => new("Themed Style 2 - Accent 4") { Guid = "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}" };

    internal static TableStyle ThemedStyle2Accent5 => new("Themed Style 2 - Accent 5") { Guid = "{327F97BB-C833-4FB7-BDE5-3F7075034690}" };

    internal static TableStyle ThemedStyle2Accent6 => new("Themed Style 2 - Accent 6") { Guid = "{638B1855-1B75-4FBE-930C-398BA8C253C6}" };

    internal static TableStyle LightStyle1 => new("Light Style 1") { Guid = "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}" };

    internal static TableStyle LightStyle1Accent1 => new("Light Style 1 - Accent 1") { Guid = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}" };

    internal static TableStyle LightStyle1Accent2 => new("Light Style 1 - Accent 2") { Guid = "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}" };

    internal static TableStyle LightStyle1Accent3 => new("Light Style 1 - Accent 3") { Guid = "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}" };

    internal static TableStyle LightStyle1Accent4 => new("Light Style 1 - Accent 4") { Guid = "{D27102A9-8310-4765-A935-A1911B00CA55}" };

    internal static TableStyle LightStyle1Accent5 => new("Light Style 1 - Accent 5") { Guid = "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}" };

    internal static TableStyle LightStyle1Accent6 => new("Light Style 1 - Accent 6") { Guid = "{68D230F3-CF80-4859-8CE7-A43EE81993B5}" };

    internal static TableStyle LightStyle2 => new("Light Style 2") { Guid = "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}" };

    internal static TableStyle LightStyle2Accent1 => new("Light Style 2 - Accent 1") { Guid = "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}" };

    internal static TableStyle LightStyle2Accent2 => new("Light Style 2 - Accent 2") { Guid = "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}" };

    internal static TableStyle LightStyle2Accent3 => new("Light Style 2 - Accent 3") { Guid = "{F2DE63D5-997A-4646-A377-4702673A728D}" };

    internal static TableStyle LightStyle2Accent4 => new("Light Style 2 - Accent 4") { Guid = "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}" };

    internal static TableStyle LightStyle2Accent5 => new("Light Style 2 - Accent 5") { Guid = "{5A111915-BE36-4E01-A7E5-04B1672EAD32}" };

    internal static TableStyle LightStyle2Accent6 => new("Light Style 2 - Accent 6") { Guid = "{912C8C85-51F0-491E-9774-3900AFEF0FD7}" };

    internal static TableStyle LightStyle3 => new("Light Style 3") { Guid = "{616DA210-FB5B-4158-B5E0-FEB733F419BA}" };

    internal static TableStyle LightStyle3Accent1 => new("Light Style 3 - Accent 1") { Guid = "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}" };

    internal static TableStyle LightStyle3Accent2 => new("Light Style 3 - Accent 2") { Guid = "{5DA37D80-6434-44D0-A028-1B22A696006F}" };

    internal static TableStyle LightStyle3Accent3 => new("Light Style 3 - Accent 3") { Guid = "{8799B23B-EC83-4686-B30A-512413B5E67A}" };

    internal static TableStyle LightStyle3Accent4 => new("Light Style 3 - Accent 4") { Guid = "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}" };

    internal static TableStyle LightStyle3Accent5 => new("Light Style 3 - Accent 5") { Guid = "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}" };

    internal static TableStyle LightStyle3Accent6 => new("Light Style 3 - Accent 6") { Guid = "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}" };

    internal static TableStyle MediumStyle1 => new("Medium Style 1") { Guid = "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}" };

    internal static TableStyle MediumStyle1Accent1 => new("Medium Style 1 - Accent 1") { Guid = "{B301B821-A1FF-4177-AEE7-76D212191A09}" };

    internal static TableStyle MediumStyle1Accent2 => new("Medium Style 1 - Accent 2") { Guid = "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}" };

    internal static TableStyle MediumStyle1Accent3 => new("Medium Style 1 - Accent 3") { Guid = "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}" };

    internal static TableStyle MediumStyle1Accent4 => new("Medium Style 1 - Accent 4") { Guid = "{1E171933-4619-4E11-9A3F-F7608DF75F80}" };

    internal static TableStyle MediumStyle1Accent5 => new("Medium Style 1 - Accent 5") { Guid = "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}" };

    internal static TableStyle MediumStyle1Accent6 => new("Medium Style 1 - Accent 6") { Guid = "{10A1B5D5-9B99-4C35-A422-299274C87663}" };

    internal static TableStyle MediumStyle2 => new("Medium Style 2") { Guid = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}" };

    internal static TableStyle MediumStyle2Accent1 => new("Medium Style 2 - Accent 1") { Guid = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };

    internal static TableStyle MediumStyle2Accent2 => new("Medium Style 2 - Accent 2") { Guid = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}" };

    internal static TableStyle MediumStyle2Accent3 => new("Medium Style 2 - Accent 3") { Guid = "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}" };

    internal static TableStyle MediumStyle2Accent4 => new("Medium Style 2 - Accent 4") { Guid = "{00A15C55-8517-42AA-B614-E9B94910E393}" };

    internal static TableStyle MediumStyle2Accent5 => new("Medium Style 2 - Accent 5") { Guid = "{7DF18680-E054-41AD-8BC1-D1AEF772440D}" };

    internal static TableStyle MediumStyle2Accent6 => new("Medium Style 2 - Accent 6") { Guid = "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}" };

    internal static TableStyle MediumStyle3 => new("Medium Style 3") { Guid = "{8EC20E35-A176-4012-BC5E-935CFFF8708E}" };

    internal static TableStyle MediumStyle3Accent1 => new("Medium Style 3 - Accent 1") { Guid = "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}" };

    internal static TableStyle MediumStyle3Accent2 => new("Medium Style 3 - Accent 2") { Guid = "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}" };

    internal static TableStyle MediumStyle3Accent3 => new("Medium Style 3 - Accent 3") { Guid = "{EB344D84-9AFB-497E-A393-DC336BA19D2E}" };

    internal static TableStyle MediumStyle3Accent4 => new("Medium Style 3 - Accent 4") { Guid = "{EB9631B5-78F2-41C9-869B-9F39066F8104}" };

    internal static TableStyle MediumStyle3Accent5 => new("Medium Style 3 - Accent 5") { Guid = "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}" };

    internal static TableStyle MediumStyle3Accent6 => new("Medium Style 3 - Accent 6") { Guid = "{2A488322-F2BA-4B5B-9748-0D474271808F}" };

    internal static TableStyle MediumStyle4 => new("Medium Style 4") { Guid = "{D7AC3CCA-C797-4891-BE02-D94E43425B78}" };

    internal static TableStyle MediumStyle4Accent1 => new("Medium Style 4 - Accent 1") { Guid = "{69CF1AB2-1976-4502-BF36-3FF5EA218861}" };

    internal static TableStyle MediumStyle4Accent2 => new("Medium Style 4 - Accent 2") { Guid = "{8A107856-5554-42FB-B03E-39F5DBC370BA}" };

    internal static TableStyle MediumStyle4Accent3 => new("Medium Style 4 - Accent 3") { Guid = "{0505E3EF-67EA-436B-97B2-0124C06EBD24}" };

    internal static TableStyle MediumStyle4Accent4 => new("Medium Style 4 - Accent 4") { Guid = "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}" };

    internal static TableStyle MediumStyle4Accent5 => new("Medium Style 4 - Accent 5") { Guid = "{22838BEF-8BB2-4498-84A7-C5851F593DF1}" };

    internal static TableStyle MediumStyle4Accent6 => new("Medium Style 4 - Accent 6") { Guid = "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}" };

    internal static TableStyle DarkStyle1 => new("Dark Style 1") { Guid = "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}" };

    internal static TableStyle DarkStyle1Accent1 => new("Dark Style 1 - Accent 1") { Guid = "{125E5076-3810-47DD-B79F-674D7AD40C01}" };

    internal static TableStyle DarkStyle1Accent2 => new("Dark Style 1 - Accent 2") { Guid = "{37CE84F3-28C3-443E-9E96-99CF82512B78}" };

    internal static TableStyle DarkStyle1Accent3 => new("Dark Style 1 - Accent 3") { Guid = "{D03447BB-5D67-496B-8E87-E561075AD55C}" };

    internal static TableStyle DarkStyle1Accent4 => new("Dark Style 1 - Accent 4") { Guid = "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}" };

    internal static TableStyle DarkStyle1Accent5 => new("Dark Style 1 - Accent 5") { Guid = "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}" };

    internal static TableStyle DarkStyle1Accent6 => new("Dark Style 1 - Accent 6") { Guid = "{AF606853-7671-496A-8E4F-DF71F8EC918B}" };

    internal static TableStyle DarkStyle2 => new("Dark Style 2") { Guid = "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}" };

    internal static TableStyle DarkStyle2Accent1Accent2 => new("Dark Style 2 - Accent 1, Accent 2") { Guid = "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}" };

    internal static TableStyle DarkStyle2Accent3Accent4 => new("Dark Style 2 - Accent 3, Accent 4") { Guid = "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}" };

    internal static TableStyle DarkStyle2Accent5Accent6 => new("Dark Style 2 - Accent 5, Accent 6") { Guid = "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}" };

    public override bool Equals(object? obj)
    {
        return obj is TableStyle style &&
               this.Name == style.Name &&
               this.Guid == style.Guid;
    }

    public override int GetHashCode()
    {
        int hashCode = 1242478914;
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(this.Name);
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(this.Guid);
        return hashCode;
    }
}

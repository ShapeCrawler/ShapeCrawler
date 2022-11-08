using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class TextCharacterPropertiesTypeExtensions
{
    internal static A.SolidFill? GetASolidFill(this A.TextCharacterPropertiesType aTextCharPropertyType)
    {
        return aTextCharPropertyType.GetFirstChild<A.SolidFill>();
    }
}
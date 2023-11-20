using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class TextCharacterPropertiesTypeExtensions
{
    internal static A.SolidFill? SDKASolidFill(this A.TextCharacterPropertiesType aTextCharPropertyType)
    {
        return aTextCharPropertyType.GetFirstChild<A.SolidFill>();
    }
}
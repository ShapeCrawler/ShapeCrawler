using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions
{
    internal static class TextCharacterPropertiesTypeExtensions
    {
        public static A.SolidFill SolidFill(this A.TextCharacterPropertiesType aTextCharPropertyType)
        {
            return aTextCharPropertyType.GetFirstChild<A.SolidFill>();
        }
    }
}
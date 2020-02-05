using SlideXML.Enums;

namespace SlideXML.Services.Placeholders
{
    public class PlaceholderXML //TODO: add Compareable
    {
        public PlaceholderType PlaceholderType { get; set; }

        /// <summary>
        /// Gets or sets index (p:ph idx="12345").
        /// </summary>
        /// <returns>Index value or null if such index not exist.</returns>
        public int? Index { get; set; }
    }
}

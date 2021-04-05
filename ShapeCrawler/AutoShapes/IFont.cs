namespace ShapeCrawler.AutoShapes
{
    public interface IFont //TODO: consider moving font properties on Portion level
    {
        /// <summary>
        ///     Gets font name.
        /// </summary>
        string Name { get; set; }

        /// <summary>
        ///     Gets or sets font size in EMUs.
        /// </summary>
        int Size { get; set; } //TODO: create test to verify font size of table cell's text portion

        /// <summary>
        ///     Gets or sets flag indicating whether font is bold.
        /// </summary>
        bool IsBold { get; set; }

        bool IsItalic { get; set; }

        /// <summary>
        ///     Gets or sets font color in HEX format.
        /// </summary>
        string Color { get; set; }

        /// <summary>
        ///     Gets value indicating whether font size can be changed.
        /// </summary>
        bool SizeCanBeChanged();
    }
}
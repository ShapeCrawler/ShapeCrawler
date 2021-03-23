namespace ShapeCrawler.AutoShapes
{
    public interface IFont
    {
        /// <summary>
        ///     Gets font name.
        /// </summary>
        string Name { get; set; }

        /// <summary>
        ///     Gets or sets font size in EMUs.
        /// </summary>
        int Size { get; set; }

        /// <summary>
        ///     Gets or sets flag indicating whether font is bold.
        /// </summary>
        bool IsBold { get; set; }

        bool IsItalic { get; set; }

        /// <summary>
        ///     Gets value indicating whether font size can be changed.
        /// </summary>
        bool SizeCanBeChanged();
    }
}
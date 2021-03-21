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

#if DEBUG
        /// <summary>
        ///     Gets or sets flag indicating whether font is bold.
        /// </summary>
        bool IsBold { get; set; }
#else
        bool IsBold { get; }
#endif

        /// <summary>
        ///     Gets value indicating whether font size can be changed.
        /// </summary>
        bool SizeCanBeChanged();
    }
}
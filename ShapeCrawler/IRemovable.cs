namespace ShapeCrawler
{
    /// <summary>
    ///     Represents interface of removable presentation's element.
    /// </summary>
    public interface IRemovable
    {
        /// <summary>
        ///     Gets or sets a value indicating whether element was removed.
        /// </summary>
        bool IsRemoved { get; set; } // TODO: make internal setter

        /// <summary>
        ///     Throws exception if element was removed.
        /// </summary>
        void ThrowIfRemoved();
    }
}
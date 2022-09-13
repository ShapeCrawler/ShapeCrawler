namespace ShapeCrawler.Services
{
    /// <summary>
    ///     Represents interface of removable presentation's element.
    /// </summary>
    internal interface IRemovable
    {
        /// <summary>
        ///     Gets or sets a value indicating whether element was removed.
        /// </summary>
        bool IsRemoved { get; set; }

        /// <summary>
        ///     Throws exception if element was removed.
        /// </summary>
        void ThrowIfRemoved();
    }
}
using System;
using System.Globalization;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown when presentation size is more than allowable.
    /// </summary>
    internal class PresentationIsLargeException : ShapeCrawlerException
    {
        private PresentationIsLargeException(string message)
            : base(message, (int)ExceptionCode.PresentationIsLargeException)
        {
        }

        /// <summary>
        ///     Creates a new instance of the <see cref="PresentationIsLargeException" /> class with specifying max presentation
        ///     size.
        /// </summary>
        internal static PresentationIsLargeException FromMax(int maxSize)
        {
#if NETSTANDARD2_1 || NET5_0 || NETCOREAPP2_1
            var message = ExceptionMessages.PresentationIsLarge.Replace("{0}",
                maxSize.ToString(CultureInfo.CurrentCulture), StringComparison.OrdinalIgnoreCase);
#else
            var message =
                ExceptionMessages.PresentationIsLarge.Replace("{0}", maxSize.ToString(CultureInfo.CurrentCulture));
#endif
            return new PresentationIsLargeException(message);
        }
    }
}
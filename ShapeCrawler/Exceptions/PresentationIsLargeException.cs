using System;
using ShapeCrawler.Enums;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    /// Thrown when presentation size is more than allowable.
    /// </summary>
    public class PresentationIsLargeException : ShapeCrawlerException
    {
        #region Constructors

        private PresentationIsLargeException(string message) : base(message, (int)ExceptionCode.PresentationIsLargeException)
        {

        }

        #endregion Constructors

        /// <summary>
        /// Creates a new instance of the <see cref="PresentationIsLargeException"/> class with specifying max presentation size.
        /// </summary>
        /// <param name="maxSize"></param>
        public static PresentationIsLargeException FromMax(int maxSize)
        {
#if NETSTANDARD2_1 || NETCOREAPP2_0
            var message = ExceptionMessages.PresentationIsLarge.Replace("{0}", maxSize.ToString(), StringComparison.OrdinalIgnoreCase);
#else
            var message = ExceptionMessages.PresentationIsLarge.Replace("{0}", maxSize.ToString());
#endif
            return new PresentationIsLargeException(message);
        }

        public PresentationIsLargeException()
        {
        }

        public PresentationIsLargeException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

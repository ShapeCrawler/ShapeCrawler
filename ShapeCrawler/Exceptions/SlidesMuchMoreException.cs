using System;
using System.Globalization;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown when number of slides more than allowed.
    /// </summary>
    internal class SlidesMuchMoreException : ShapeCrawlerException
    {
        #region Constructors

        private SlidesMuchMoreException(string message) : base(message, (int) ExceptionCode.SlidesMuchMoreException)
        {
        }

        #endregion Constructors

        internal SlidesMuchMoreException()
        {
        }

        public SlidesMuchMoreException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public static SlidesMuchMoreException FromMax(int maxNum)
        {
#if NETSTANDARD2_1 || NET5_0 || NETCOREAPP2_1
            var message = ExceptionMessages.SlidesMuchMore.Replace("{0}", maxNum.ToString(CultureInfo.CurrentCulture),
                StringComparison.OrdinalIgnoreCase);
#else
            var message = ExceptionMessages.SlidesMuchMore.Replace("{0}", maxNum.ToString(CultureInfo.CurrentCulture));
#endif
            return new SlidesMuchMoreException(message);
        }
    }
}
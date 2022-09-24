using System;
using System.Globalization;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown when number of slides more than allowed.
    /// </summary>
    internal class SlidesMuchMoreException : ShapeCrawlerException
    {
        private SlidesMuchMoreException(string message)
            : base(message, (int) ExceptionCode.SlidesMuchMoreException)
        {
        }

        internal static SlidesMuchMoreException FromMax(int maxNum)
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
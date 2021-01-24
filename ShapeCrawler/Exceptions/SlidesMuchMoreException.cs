using System;
using System.Globalization;
using ShapeCrawler.Enums;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    /// Thrown when number of slides more than allowed.
    /// </summary>
    public class SlidesMuchMoreException : ShapeCrawlerException
    {
        #region Constructors

        private SlidesMuchMoreException(string message) : base(message, (int)ExceptionCode.SlidesMuchMoreException) { }

        #endregion Constructors

        public static SlidesMuchMoreException FromMax(int maxNum)
        {
#if NETSTANDARD2_1 || NETCOREAPP2_0
            var message = ExceptionMessages.SlidesMuchMore.Replace("{0}", maxNum.ToString(CultureInfo.CurrentCulture), StringComparison.OrdinalIgnoreCase);
#else
            var message = ExceptionMessages.SlidesMuchMore.Replace("{0}", maxNum.ToString(CultureInfo.CurrentCulture));
#endif
            return new SlidesMuchMoreException(message);
        }

        public SlidesMuchMoreException()
        {
        }

        public SlidesMuchMoreException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
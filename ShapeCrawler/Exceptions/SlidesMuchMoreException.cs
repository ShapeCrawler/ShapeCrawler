using System.Globalization;

namespace ShapeCrawler.Exceptions;

internal sealed class SlidesMuchMoreException : ShapeCrawlerException
{
    private SlidesMuchMoreException(string message)
        : base(message, (int)ExceptionCode.SlidesMuchMoreException)
    {
    }

    internal static SlidesMuchMoreException FromMax(int maxNum)
    {
#if NET7_0
        var message = ExceptionMessages.SlidesMuchMore.Replace(
            "{0}", 
            maxNum.ToString(CultureInfo.CurrentCulture),
            System.StringComparison.OrdinalIgnoreCase);
#else
            var message = ExceptionMessages.SlidesMuchMore.Replace("{0}", maxNum.ToString(CultureInfo.CurrentCulture));
#endif
        return new SlidesMuchMoreException(message);
    }
}
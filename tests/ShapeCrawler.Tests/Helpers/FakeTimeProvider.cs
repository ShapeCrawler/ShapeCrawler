using ShapeCrawler.Presentations;

namespace ShapeCrawler.Tests.Helpers;

internal class FakeTimeProvider(DateTime date): ITimeProvider
{
    DateTime ITimeProvider.UtcNow => date;
}

using ShapeCrawler.Presentations;

namespace ShapeCrawler.DevTests.Helpers;

internal class FakeTimeProvider(DateTime date): ITimeProvider
{
    DateTime ITimeProvider.UtcNow => date;
}

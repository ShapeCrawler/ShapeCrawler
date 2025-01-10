using ShapeCrawler.Shared;

namespace ShapeCrawler.Tests.Unit.Helpers;

internal class FakeTimeProvider(DateTime date): ITimeProvider
{
    DateTime ITimeProvider.UtcNow => date;
}

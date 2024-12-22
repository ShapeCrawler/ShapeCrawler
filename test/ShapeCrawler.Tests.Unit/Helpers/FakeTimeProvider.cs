using ShapeCrawler.Shared;

namespace ShapeCrawler.Tests.Unit.Helpers;

/// <summary>
/// Provides a faked time which can be controlled by unit tests.
/// </summary>
/// <param name="fakeTime">Fake time to return when asked.</param>
internal class FakeTimeProvider(DateTime fakeTime): ITimeProvider
{
    DateTime ITimeProvider.UtcNow => fakeTime;
}

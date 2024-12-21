using System;

namespace ShapeCrawler.Shared;

/// <summary>
/// Provides the current date and time.
/// </summary>
internal interface ITimeProvider
{
    /// <summary>
    /// Gets current date and time.
    /// </summary>
    DateTime UtcNow { get; }
}

/// <summary>
/// Provides the actual real current date and time.
/// </summary>
internal class SystemTimeProvider: ITimeProvider
{
    DateTime ITimeProvider.UtcNow => DateTime.UtcNow;
}

/// <summary>
/// Provides a faked time which can be controlled by unit tests.
/// </summary>
/// <param name="fakeTime">Fake time to return when asked.</param>
internal class FakeTimeProvider(DateTime fakeTime): ITimeProvider
{
    DateTime ITimeProvider.UtcNow => fakeTime;
}

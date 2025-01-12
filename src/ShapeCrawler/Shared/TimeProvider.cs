using System;

namespace ShapeCrawler.Shared;

/// <summary>
///     Provides the current date and time.
/// </summary>
internal interface ITimeProvider
{
    /// <summary>
    ///     Gets the current date and time.
    /// </summary>
    DateTime UtcNow { get; }
}

internal class SystemTimeProvider: ITimeProvider
{
    DateTime ITimeProvider.UtcNow => DateTime.UtcNow;
}

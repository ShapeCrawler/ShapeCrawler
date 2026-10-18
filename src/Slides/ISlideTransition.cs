#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Controls how a slide advances during a slide show.
/// </summary>
public interface ISlideTransition
{
    /// <summary>
    ///     Gets or sets a value indicating whether a mouse click advances the slide.
    ///     The default is <see langword="true"/>.
    /// </summary>
    bool AdvanceOnClick { get; set; }

    /// <summary>
    ///     Gets or sets the delay in milliseconds before automatically advancing the slide.
    ///     A value of <see langword="null"/> disables automatic advance (the default).
    ///     Zero enables immediate advance. The maximum is 2147483647 milliseconds.
    ///     Click advance is controlled independently.
    /// </summary>
    uint? AdvanceAfterTime { get; set; }
}

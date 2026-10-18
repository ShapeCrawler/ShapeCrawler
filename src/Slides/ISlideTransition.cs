#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Controls how a slide advances during a slide show.
/// </summary>
public interface ISlideTransition
{
    /// <summary>
    ///     Gets or sets the effect. The default is <see cref="SlideTransitionType.None"/>.
    ///     Existing newer effects return <see cref="SlideTransitionType.Unsupported"/>, which cannot be assigned.
    ///     Changing the type uses the new effect's schema defaults and clears any explicit custom duration.
    ///     Assigning the same type preserves its existing parameters.
    /// </summary>
    SlideTransitionType Type { get; set; }

    /// <summary>
    ///     Gets or sets the transition speed. The default is <see cref="SlideTransitionSpeed.Fast"/>.
    ///     Setting the speed clears any explicit custom duration so that the selected speed takes effect.
    /// </summary>
    SlideTransitionSpeed Speed { get; set; }

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

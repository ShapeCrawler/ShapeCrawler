#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Audio start mode.
/// </summary>
public enum AudioStartMode
{
    /// <summary>
    ///     In Click Sequence.
    /// </summary>
    InClickSequence,

    /// <summary>
    ///     Automatically (With Previous in the animation Timing options).
    /// </summary>
    Automatically,

    /// <summary>
    ///     When Clicked On.
    /// </summary>
    WhenClickedOn
}

using System.Collections.Generic;
using ShapeCrawler.SlideMasters;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Slide Master.
/// </summary>
public interface ISlideMaster
{
    /// <summary>
    ///     Gets background image if slide master has background, otherwise <see langword="null"/>.
    /// </summary>
    IImage? Background { get; }

    /// <summary>
    ///     Gets collection of Slide Layouts.
    /// </summary>
    IReadOnlyList<ISlideLayout> SlideLayouts { get; }

    /// <summary>
    ///     Gets collection of shape.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets parent Presentation
    /// </summary>
    IPresentation Presentation { get; }

    /// <summary>
    ///     Gets theme.
    /// </summary>
    ITheme Theme { get; }
}

/// <summary>
///     Represents a theme.
/// </summary>
public interface ITheme
{
    /// <summary>
    ///     Gets font settings.
    /// </summary>
    IThemeFontSetting FontSettings { get; }
}

/// <summary>
///     Represents a settings of theme font.
/// </summary>
public interface IThemeFontSetting
{
    /// <summary>
    ///     Gets heading font name.
    /// </summary>
    string Head { get; set; }

    /// <summary>
    ///     Gets body font name.
    /// </summary>
    string Body { get; set; }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents color scheme resolution for shapes.
/// </summary>
internal sealed class ShapeColorScheme
{
    private readonly OpenXmlElement pShapeTreeElement;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ShapeColorScheme"/> class.
    /// </summary>
    /// <param name="pShapeTreeElement">The shape tree element.</param>
    internal ShapeColorScheme(OpenXmlElement pShapeTreeElement)
    {
        this.pShapeTreeElement = pShapeTreeElement;
    }

    /// <summary>
    ///     Gets the color scheme.
    /// </summary>
    /// <returns>The color scheme.</returns>
    internal A.ColorScheme? GetColorScheme()
    {
        var parentPart = new SCOpenXmlElement(this.pShapeTreeElement).ParentOpenXmlPart;

        return parentPart switch
        {
            SlidePart slidePart => GetColorSchemeFromSlidePart(slidePart),
            SlideLayoutPart slideLayoutPart => GetColorSchemeFromSlideLayoutPart(slideLayoutPart),
            SlideMasterPart slideMasterPart => GetColorSchemeFromSlideMasterPart(slideMasterPart),
            NotesSlidePart notesSlidePart => GetColorSchemeFromNotesSlidePart(notesSlidePart),
            _ => null
        };
    }

    private static A.ColorScheme? GetColorSchemeFromSlidePart(SlidePart slidePart)
    {
        var slideLayoutPart = slidePart.SlideLayoutPart;
        if (slideLayoutPart is null)
        {
            return null;
        }

        return GetColorSchemeFromSlideLayoutPart(slideLayoutPart);
    }

    private static A.ColorScheme? GetColorSchemeFromSlideLayoutPart(SlideLayoutPart slideLayoutPart)
    {
        var slideMasterPart = slideLayoutPart.SlideMasterPart;
        if (slideMasterPart is null)
        {
            return null;
        }

        return GetColorSchemeFromSlideMasterPart(slideMasterPart);
    }

    private static A.ColorScheme? GetColorSchemeFromSlideMasterPart(SlideMasterPart slideMasterPart)
    {
        var themePart = slideMasterPart.ThemePart;
        var themeElements = themePart?.Theme.ThemeElements;

        return themeElements?.ColorScheme;
    }

    private static A.ColorScheme? GetColorSchemeFromNotesSlidePart(NotesSlidePart notesSlidePart)
    {
        var notesMasterPart = notesSlidePart.NotesMasterPart;
        if (notesMasterPart is null)
        {
            return null;
        }

        var themePart = notesMasterPart.ThemePart;
        var themeElements = themePart?.Theme.ThemeElements;

        return themeElements?.ColorScheme;
    }
}

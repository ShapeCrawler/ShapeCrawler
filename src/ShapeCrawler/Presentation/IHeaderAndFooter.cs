using System.Linq;

namespace ShapeCrawler;

/// <summary>
///     Represents Header and Footer manager.
/// </summary>
public interface IHeaderAndFooter
{
    /// <summary>
    ///     Gets a value indicating whether slide number is visible.
    /// </summary>
    bool SlideNumberAdded();

    /// <summary>
    ///     Adds slide number on slides.
    /// </summary>
    void AddSlideNumber();

    /// <summary>
    ///     Removes slide number from slides.
    /// </summary>
    void RemoveSlideNumber();
}

internal sealed class HeaderAndFooter : IHeaderAndFooter
{
    private readonly Presentation presentation;

    internal HeaderAndFooter(Presentation presentation)
    {
        this.presentation = presentation;
    }

    public bool SlideNumberAdded()
    {
        return this.presentation.Slides.Any(slide =>
            slide.Shapes.Any(shape => shape.Placeholder?.Type == SCPlaceholderType.SlideNumber));
    }

    public void AddSlideNumber()
    {
        if (this.SlideNumberAdded())
        {
            return;
        }

        foreach (var slide in this.presentation.Slides)
        {
            var slideNumberPlaceholder =
                slide.SlideLayout.Shapes.FirstOrDefault(shape =>
                    shape.Placeholder?.Type == SCPlaceholderType.SlideNumber);
            if (slideNumberPlaceholder != null)
            {
                slide.Shapes.Add(slideNumberPlaceholder);
            }
        }
    }

    public void RemoveSlideNumber()
    {
        if (!this.SlideNumberAdded())
        {
            return;
        }

        foreach (var slide in this.presentation.Slides)
        {
            var slideNumberPlaceholder =
                slide.Shapes.FirstOrDefault(shape =>
                    shape.Placeholder?.Type == SCPlaceholderType.SlideNumber);
            if (slideNumberPlaceholder != null)
            {
                slide.Shapes.Remove(slideNumberPlaceholder);
            }
        }
    }
}
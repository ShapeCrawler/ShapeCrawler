using System.Linq;

namespace ShapeCrawler;

/// <summary>
///     Represents Header and Footer manager.
/// </summary>
public interface IHeaderFooterManager
{
    /// <summary>
    ///     Gets a value indicating whether slide number is visible.
    /// </summary>
    bool IsSlideNumberVisible();

    /// <summary>
    ///     Sets slide number visibility.
    /// </summary>
    void SetSlideNumberVisible(bool visible);
}

internal class HeaderFooterManager : IHeaderFooterManager
{
    private readonly SCPresentation presentation;

    internal HeaderFooterManager(SCPresentation presentation)
    {
        this.presentation = presentation;
    }

    public bool IsSlideNumberVisible()
    {
        return this.presentation.Slides.Any(slide => slide.Shapes.Any(shape => shape.Placeholder?.Type == SCPlaceholderType.SlideNumber));
    }

    public void SetSlideNumberVisible(bool visible)
    {
        if (this.IsSlideNumberVisible() == visible)
        {
            return;
        }

        if (visible)
        {
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
        else
        {
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
}
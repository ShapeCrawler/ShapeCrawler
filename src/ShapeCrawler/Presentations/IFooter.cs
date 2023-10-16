using System.Linq;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler;

/// <summary>
///     Represents Header and Footer manager.
/// </summary>
public interface IFooter
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

internal sealed class Footer : IFooter
{
    private readonly PresentationCore presentationCore;

    internal Footer(PresentationCore presentationCore)
    {
        this.presentationCore = presentationCore;
    }

    public bool SlideNumberAdded() 
    {
        return this.presentationCore.Slides.Any(slide =>
            slide.Shapes.Any(shape => shape.PlaceholderType == PlaceholderType.SlideNumber));
    }

    public void AddSlideNumber()
    {
        if (this.SlideNumberAdded())
        {
            return;
        }

        foreach (var slide in this.presentationCore.Slides)
        {
            var slideNumberPlaceholder =
                slide.SlideLayout.Shapes.FirstOrDefault(shape =>
                    shape.PlaceholderType == PlaceholderType.SlideNumber);
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

        foreach (var slide in this.presentationCore.Slides)
        {
            var slideNumberPlaceholder =
                slide.Shapes.FirstOrDefault(shape =>
                    shape.PlaceholderType == PlaceholderType.SlideNumber);
            if (slideNumberPlaceholder != null)
            {
                slide.Shapes.Remove(slideNumberPlaceholder);
            }
        }
    }
}
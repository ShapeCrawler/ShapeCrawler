using System.Linq;
using ShapeCrawler.Slides;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents Footer.
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

internal sealed class Footer(SlideCollection slides): IFooter
{
    public bool SlideNumberAdded() 
    {
        return slides.Any(slide =>
            slide.Shapes.Any(shape => shape.PlaceholderType == PlaceholderType.SlideNumber));
    }

    public void AddSlideNumber()
    {
        if (this.SlideNumberAdded())
        {
            return;
        }

        foreach (var slide in slides)
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

        foreach (var slide in slides)
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
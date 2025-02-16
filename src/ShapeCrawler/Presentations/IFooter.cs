using System.Linq;
using ShapeCrawler.Presentations;

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

internal sealed class Footer : IFooter
{
    private readonly PresentationCore presentationCore;

    internal Footer(PresentationCore presentationCore)
    {
        this.presentationCore = presentationCore;
    }

    public bool SlideNumberAdded() 
    {
        return this.presentationCore.SlideCollection.Any(slide =>
            slide.ShapeCollection.Any(shape => shape.PlaceholderType == PlaceholderType.SlideNumber));
    }

    public void AddSlideNumber()
    {
        if (this.SlideNumberAdded())
        {
            return;
        }

        foreach (var slide in this.presentationCore.SlideCollection)
        {
            var slideNumberPlaceholder =
                slide.SlideLayout.Shapes.FirstOrDefault(shape =>
                    shape.PlaceholderType == PlaceholderType.SlideNumber);
            if (slideNumberPlaceholder != null)
            {
                slide.ShapeCollection.Add(slideNumberPlaceholder);
            }
        }
    }

    public void RemoveSlideNumber()
    {
        if (!this.SlideNumberAdded())
        {
            return;
        }

        foreach (var slide in this.presentationCore.SlideCollection)
        {
            var slideNumberPlaceholder =
                slide.ShapeCollection.FirstOrDefault(shape =>
                    shape.PlaceholderType == PlaceholderType.SlideNumber);
            if (slideNumberPlaceholder != null)
            {
                slide.ShapeCollection.Remove(slideNumberPlaceholder);
            }
        }
    }
}
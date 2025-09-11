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

    /// <summary>
    ///     Set footer text on all slides.
    /// </summary>
    void AddFooterText(string text);

    /// <summary>
    ///     Removes footer text from all slides.
    /// </summary>
    void RemoveFooterText();

    /// <summary>
    ///     Removes footer shape from all slides.
    /// </summary>
    void RemoveFooter();
}

internal sealed class Footer(UpdatedSlideCollection slides): IFooter
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
            slideNumberPlaceholder?.Remove();
        }
    }

    public void AddFooterText(string text)
    {
        foreach (var slide in slides)
        {
            var footerShape = slide.Shapes.FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);
            if (footerShape != null && footerShape.TextBox != null)
            {
                footerShape.TextBox.SetText(text);
            }
            else
            {
                var layoutFooterShape = slide.SlideLayout.Shapes.FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);
                if (layoutFooterShape != null)
                {
                    layoutFooterShape.TextBox?.SetText(text);
                    slide.Shapes.Add(layoutFooterShape);
                }
            }
        }
    }

    public void RemoveFooterText()
    {
        foreach (var slide in slides)
        {
            var footerShape = slide.Shapes.FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);
            footerShape?.TextBox?.SetText(string.Empty);
        }
    }

    public void RemoveFooter()
    {
        foreach (var slide in slides)
        {
            var footerShape = slide.Shapes.FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);
            footerShape?.Remove();
        }
    }
}
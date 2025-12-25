using System;
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
    ///     Adds text footer on all slides with specified content. If a slide already has a text footer, the content will be replaced.
    /// </summary>
    /// <param name="text">Text content.</param>
    void AddText(string text);

    /// <summary>
    ///     Removes text footers from all slides.
    /// </summary>
    void RemoveText();

    /// <summary>
    ///     Adds text footer on a specific slide with specified content. If the slide already has a text footer, the content will be replaced.
    /// </summary>
    /// <param name="slideNumber">Slide number.</param>
    /// <param name="text">Text content.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="slideNumber"/> is less than 1 or greater than the number of slides.
    /// </exception>
    void AddTextOnSlide(int slideNumber, string text);

    /// <summary>
    ///     Removes text footer from a specific slide.
    /// </summary>
    /// <param name="slideNumber">Slide number.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="slideNumber"/> is less than 1 or greater than the number of slides.
    /// </exception>
    void RemoveTextOnSlide(int slideNumber);
}

internal sealed class Footer(UpdatedSlideCollection slides) : IFooter
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
                slide.LayoutSlide.Shapes.FirstOrDefault(shape =>
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

    public void AddText(string text)
    {
        for (var i = 1; i <= slides.Count; i++)
        {
            this.AddTextOnSlide(i, text);
        }
    }

    public void RemoveText()
    {
        for (var i = 1; i <= slides.Count; i++)
        {
            this.RemoveTextOnSlide(i);
        }
    }

    public void AddTextOnSlide(int slideNumber, string text)
    {
        if (slideNumber < 1 || slideNumber > slides.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(slideNumber));
        }

        var slide = slides[slideNumber - 1];
        var existingFooterShape = slide.Shapes.FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);

        if (existingFooterShape?.TextBox != null)
        {
            existingFooterShape.TextBox.SetText(text);
            return;
        }

        var layoutFooter = slide.LayoutSlide.Shapes
            .FirstOrDefault(s => s.PlaceholderType == PlaceholderType.Footer);

        if (layoutFooter?.TextBox != null)
        {
            layoutFooter.TextBox.SetText(text);
            slide.Shapes.Add(layoutFooter);
        }
    }

    public void RemoveTextOnSlide(int slideNumber)
    {
        if (slideNumber < 1 || slideNumber > slides.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(slideNumber));
        }

        var footerShape = slides[slideNumber - 1].Shapes.FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);
        footerShape?.Remove();
    }
}
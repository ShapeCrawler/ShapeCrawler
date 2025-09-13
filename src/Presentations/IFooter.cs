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
    void AddText(string text);

    /// <summary>
    ///     Removes text footers from all slides.
    /// </summary>
    void RemoveText();
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

    public void AddText(string text)
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

    public void RemoveText()
    {
        foreach (var slide in slides)
        {
            var footerShape = slide.Shapes.FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);
            footerShape?.Remove();
        }
    }
}
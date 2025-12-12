using System;
using System.Collections.Generic;
using System.Linq;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft for building a presentation with a fluent API.
/// </summary>
public sealed class DraftPresentation
{
    private readonly List<Action<Presentation>> actions = [];
    private readonly Presentation? presentation;

    internal DraftPresentation()
    {
    }

    internal DraftPresentation(Presentation presentation)
    {
        this.presentation = presentation;
    }

    /// <summary>
    ///     Configures a slide within the presentation draft.
    ///     For a new presentation this targets the first slide.
    /// </summary>
    public DraftPresentation Slide()
    {
        var slideDraft = new DraftSlide();
        this.actions.Add(p => slideDraft.ApplyTo(p));
        return this;
    }

    /// <summary>
    ///     Configures a slide within the presentation draft.
    ///     For a new presentation this targets the first slide.
    /// </summary>
    public DraftPresentation Slide(Action<DraftSlide> configure)
    {
        var slideDraft = new DraftSlide();
        configure(slideDraft);
        this.actions.Add(p => slideDraft.ApplyTo(p));
        return this;
    }

    /// <summary>
    ///     Adds a new slide using the specified layout.
    /// </summary>
    public DraftPresentation Slide(ILayoutSlide layout)
    {
        this.actions.Add(p =>
        {
            // If no slides yet, create the initial slide first to ensure consistent numbering
            if (p.Slides.Count == 0)
            {
                var blank = p.SlideMaster(1).LayoutSlides.First(l => l.Name == "Blank");
                p.Slides.Add(blank.Number);
            }

            p.Slides.Add(layout.Number);
        });
        return this;
    }

    /// <summary>
    ///     Adds a new slide using the slide layout found by name on the first slide master.
    /// </summary>
    public DraftPresentation Slide(string layoutName)
    {
        if (string.IsNullOrWhiteSpace(layoutName))
        {
            throw new ArgumentException("Layout name must be provided", nameof(layoutName));
        }

        this.actions.Add(p =>
        {
            var layout = p.SlideMaster(1).SlideLayout(layoutName);
            if (p.Slides.Count == 0)
            {
                // Ensure initial slide id list and add via layout
            }

            p.Slides.Add(layout.Number);
        });
        return this;
    }

    /// <summary>
    ///     Gets slide master by number.
    /// </summary>
    public IMasterSlide SlideMaster(int number)
    {
        if (this.presentation == null)
        {
            throw new InvalidOperationException("Presentation has not been initialized.");
        }

        return this.presentation.SlideMaster(number);
    }

    internal void ApplyTo(Presentation paramPresentation)
    {
        foreach (var action in this.actions)
        {
            action(paramPresentation);
        }
    }
}
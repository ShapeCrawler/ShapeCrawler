using System;
using System.Collections.Generic;
using System.Linq;

namespace ShapeCrawler;

public sealed partial class Presentation
{
    /// <summary>
    ///     Represents a draft for building a presentation with a fluent API.
    /// </summary>
    public sealed class DraftPresentation
    {
        private readonly List<Action<Presentation>> actions = [];

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
        public DraftPresentation Slide(ISlideLayout layout)
        {
            this.actions.Add(p =>
            {
                // If no slides yet, create the initial slide first to ensure consistent numbering
                if (p.Slides.Count == 0)
                {
                    var blank = p.SlideMaster(1).SlideLayouts.First(l => l.Name == "Blank");
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
        
        internal void ApplyTo(Presentation presentation)
        {
            foreach (var action in this.actions)
            {
                action(presentation);
            }
        }
    }
}
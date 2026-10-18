using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideTransition(SlidePart slidePart) : ISlideTransition
{
    public bool AdvanceOnClick
    {
        get => this.Transitions().FirstOrDefault()?.AdvanceOnClick?.Value ?? true;
        set
        {
            foreach (var transition in this.ExistingOrNewTransitions())
            {
                transition.AdvanceOnClick = value;
            }
        }
    }

    public uint? AdvanceAfterTime
    {
        get
        {
            var text = this.Transitions().FirstOrDefault()?.AdvanceAfterTime?.Value;
            return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var time)
                ? time
                : null;
        }
        set
        {
            if (value > int.MaxValue)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The delay cannot exceed 2147483647 milliseconds.");
            }

            var transitions = value.HasValue ? this.ExistingOrNewTransitions() : this.Transitions();
            foreach (var transition in transitions)
            {
                if (value.HasValue)
                {
                    transition.AdvanceAfterTime = value.Value.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    transition.AdvanceAfterTime = null;
                }
            }
        }
    }

    private IEnumerable<P.Transition> Transitions() => slidePart.Slide!.Descendants<P.Transition>();

    private IEnumerable<P.Transition> ExistingOrNewTransitions()
    {
        // Update both Choice and Fallback when a transition uses markup compatibility.
        var transitions = this.Transitions().ToArray();
        if (transitions.Length > 0)
        {
            return transitions;
        }

        var transition = new P.Transition();
        slidePart.Slide!.AddChild(transition, true);
        return [transition];
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideTransition(SlidePart slidePart) : ISlideTransition
{
    private const string DurationNamespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    public SlideTransitionType Type
    {
        get => EffectType(this.Transitions().FirstOrDefault());
        set
        {
            if (!Enum.IsDefined(value) || value == SlideTransitionType.Unsupported)
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            foreach (var transition in this.ExistingOrNewTransitions())
            {
                if (EffectType(transition) == value)
                {
                    continue;
                }

                foreach (var effect in Effects(transition).ToArray())
                {
                    effect.Remove();
                }

                transition.RemoveAttribute("dur", DurationNamespace);
                var replacement = NewEffect(value);
                if (replacement is not null)
                {
                    transition.AddChild(replacement, true);
                }
            }
        }
    }

    public SlideTransitionSpeed Speed
    {
        get
        {
            var speed = this.Transitions().FirstOrDefault()?.Speed?.Value;
            if (speed == P.TransitionSpeedValues.Slow)
            {
                return SlideTransitionSpeed.Slow;
            }

            return speed == P.TransitionSpeedValues.Medium ? SlideTransitionSpeed.Medium : SlideTransitionSpeed.Fast;
        }
        set
        {
            if (!Enum.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            foreach (var transition in this.ExistingOrNewTransitions())
            {
                transition.Speed = value switch
                {
                    SlideTransitionSpeed.Slow => P.TransitionSpeedValues.Slow,
                    SlideTransitionSpeed.Medium => P.TransitionSpeedValues.Medium,
                    _ => P.TransitionSpeedValues.Fast
                };
                transition.RemoveAttribute("dur", DurationNamespace);
            }
        }
    }

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

    private static IEnumerable<OpenXmlElement> Effects(P.Transition transition) =>
        transition.ChildElements.Where(child => child is not P.SoundAction and not P.ExtensionListWithModification);

    private static SlideTransitionType EffectType(P.Transition? transition)
    {
        var effect = transition is null ? null : Effects(transition).FirstOrDefault();
        if (effect is null)
        {
            return SlideTransitionType.None;
        }

        return effect switch
        {
            P.BlindsTransition => SlideTransitionType.Blinds,
            P.CheckerTransition => SlideTransitionType.Checker,
            P.CircleTransition => SlideTransitionType.Circle,
            P.CombTransition => SlideTransitionType.Comb,
            P.CoverTransition => SlideTransitionType.Cover,
            P.CutTransition => SlideTransitionType.Cut,
            P.DiamondTransition => SlideTransitionType.Diamond,
            P.DissolveTransition => SlideTransitionType.Dissolve,
            P.FadeTransition => SlideTransitionType.Fade,
            P.NewsflashTransition => SlideTransitionType.Newsflash,
            P.PlusTransition => SlideTransitionType.Plus,
            P.PullTransition => SlideTransitionType.Pull,
            P.PushTransition => SlideTransitionType.Push,
            P.RandomTransition => SlideTransitionType.Random,
            P.RandomBarTransition => SlideTransitionType.RandomBar,
            P.SplitTransition => SlideTransitionType.Split,
            P.StripsTransition => SlideTransitionType.Strips,
            P.WedgeTransition => SlideTransitionType.Wedge,
            P.WheelTransition => SlideTransitionType.Wheel,
            P.WipeTransition => SlideTransitionType.Wipe,
            P.ZoomTransition => SlideTransitionType.Zoom,
            _ => SlideTransitionType.Unsupported
        };
    }

    private static OpenXmlElement? NewEffect(SlideTransitionType type) => type switch
    {
        SlideTransitionType.Blinds => new P.BlindsTransition(),
        SlideTransitionType.Checker => new P.CheckerTransition(),
        SlideTransitionType.Circle => new P.CircleTransition(),
        SlideTransitionType.Comb => new P.CombTransition(),
        SlideTransitionType.Cover => new P.CoverTransition(),
        SlideTransitionType.Cut => new P.CutTransition(),
        SlideTransitionType.Diamond => new P.DiamondTransition(),
        SlideTransitionType.Dissolve => new P.DissolveTransition(),
        SlideTransitionType.Fade => new P.FadeTransition(),
        SlideTransitionType.Newsflash => new P.NewsflashTransition(),
        SlideTransitionType.Plus => new P.PlusTransition(),
        SlideTransitionType.Pull => new P.PullTransition(),
        SlideTransitionType.Push => new P.PushTransition(),
        SlideTransitionType.Random => new P.RandomTransition(),
        SlideTransitionType.RandomBar => new P.RandomBarTransition(),
        SlideTransitionType.Split => new P.SplitTransition(),
        SlideTransitionType.Strips => new P.StripsTransition(),
        SlideTransitionType.Wedge => new P.WedgeTransition(),
        SlideTransitionType.Wheel => new P.WheelTransition(),
        SlideTransitionType.Wipe => new P.WipeTransition(),
        SlideTransitionType.Zoom => new P.ZoomTransition(),
        _ => null
    };

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

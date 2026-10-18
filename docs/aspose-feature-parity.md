# Aspose.Slides feature parity

This ledger tracks incremental, independently implemented ShapeCrawler features
investigated with [REA](https://github.com/morluto/rea). It records a bounded
comparison, not a claim of complete Aspose.Slides compatibility.

## Iteration rules

- Keep each PR small enough to create within 40 minutes, including investigation
  and local validation. Check all PR checks after publishing.
- Stop after each PR and wait for the maintainer's approval before the next one.
- Exclude slide-to-image rendering; it is being developed separately.
- Preserve ShapeCrawler's API conventions and keep Open XML SDK types internal.
- Record the reference version, observations, limitations, and regression tests.
  Do not copy proprietary implementation code or add an Aspose runtime dependency.

## Reference and method for iteration 1

The locally available reference is **Aspose.Slides.NET 24.12.0**, `lib/net6.0`,
with assembly SHA-256
`edaaa3005d2a462548fadcb35ffc661e2b752e7ee863c9c515f459ef7e345ca2`.
This is a versioned baseline, not an assertion about the latest Aspose release.

REA **6.7.0** was invoked through its CLI, without changing agent configuration:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-artifact $assembly --json
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.BaseSlide --type Aspose.Slides.Slide --json
```

Set `$assembly` to the reference DLL above. The first command identified an
AnyCPU, modern .NET, CIL assembly and reported complete artifact metadata
(evidence `ev_6d927c317492f227a3faaa17da3ed8a04d096e61665af76e7b970a0b6cd4c818`).
The second command returned **unavailable** member coverage: metadata exceeded
REA's 67,108,864-byte representation budget, even with type selection
(evidence `ev_a4295b8bdf10674a784e4e6b85e64ef5d6af0541a42fd7d8ad4f316ea1cd85a8`).
No method-body conclusions are drawn from that result.

The package's XML API documentation identified the relevant public properties.
A separate console probe executed the unmodified Aspose assembly in evaluation
mode, saved a PPTX, read `ppt/slides/slide1.xml`, and reopened it with Aspose.
The evaluation watermark was left intact. Only the public behavior and file
format observations below informed the ShapeCrawler implementation.

## Implemented in iteration 1

| Aspose capability | Observed reference behavior | ShapeCrawler API | PPTX representation |
| --- | --- | --- | --- |
| `BaseSlide.Name` | Initially empty; Unicode and XML-sensitive characters survive save/reopen | `IUserSlide.Name` | `p:cSld/@name` |
| `Slide.ShowMasterShapes` | Initially true; false survives save/reopen independently of visibility | `IUserSlide.ShowMasterShapes` | `p:sld/@showMasterSp` |
| Set `Slide.Hidden = false` | Hidden slide becomes visible again and remains visible after save/reopen | `IUserSlide.Show()` alongside existing `Hide()` and `Hidden()` | `p:sld/@show`; omitted or true means visible |

ShapeCrawler additionally defines `Name = null` to throw
`ArgumentNullException` without changing the name. This is its own API contract;
the reference's null behavior was not assessed. Regression tests in
`UserSlideTests` inspect saved PPTX files, cover missing attributes and state
changes, reopen the presentations, and run Open XML validation.
`PublicApiTests` continues to guard SDK encapsulation.

To repeat the reference observation in a temporary .NET console project with
`Aspose.Slides.NET` 24.12.0 installed:

```csharp
using Aspose.Slides;
using Aspose.Slides.Export;
using System.IO.Compression;

using var presentation = new Presentation();
var slide = presentation.Slides[0];
Console.WriteLine($"[{slide.Name}], {slide.Hidden}, {slide.ShowMasterShapes}");
slide.Name = "Обзор & <2026>";
slide.Hidden = true;
slide.Hidden = false;
slide.ShowMasterShapes = false;
using var stream = new MemoryStream();
presentation.Save(stream, SaveFormat.Pptx);
stream.Position = 0;
using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, true))
using (var reader = new StreamReader(zip.GetEntry("ppt/slides/slide1.xml")!.Open()))
    Console.WriteLine(reader.ReadToEnd());
stream.Position = 0;
using var reopened = new Presentation(stream);
slide = reopened.Slides[0];
Console.WriteLine($"[{slide.Name}], {slide.Hidden}, {slide.ShowMasterShapes}");
```

## Iteration 2: slide advance settings

Narrow API comparison found no ShapeCrawler equivalent for
`ISlideShowTransition.AdvanceOnClick`, `AdvanceAfter`, or `AdvanceAfterTime`.
The new `IUserSlide.Transition` returns `ISlideTransition`:

```csharp
slide.Transition.AdvanceOnClick = false;
slide.Transition.AdvanceAfterTime = 5000; // Automatic advance after 5 seconds.
slide.Transition.AdvanceAfterTime = null; // Disable automatic advance.
```

`AdvanceAfterTime` is a nullable millisecond count. Assigning a number enables
the timer; assigning null disables it. Zero enables immediate advance.
The click setting is independent. Reading the settings does not create XML.
Changing them retains existing transition effects, speed, and animation data;
markup-compatibility Choice and Fallback transitions are updated together.

### Reference evidence

The reference version and assembly hash are unchanged from iteration 1.
REA 6.7.0 was run again with the exact declaring type from the package's API
documentation:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.SlideShow.SlideShowTransition --json
```

Evidence `ev_b6e3b13b8e36a906520a0ffd9bb45a73fb13da3a631e28ae94565fcdce9de778`
again reports unavailable member coverage because of the 64 MiB metadata
representation budget. No method bodies were recovered. A separate runtime
probe used the same unmodified evaluation assembly and save/reopen procedure
as iteration 1; only the transition settings were changed.

| Reference operations on a new slide | Observed `AdvanceOnClick` / `AdvanceAfter` / `AdvanceAfterTime` after reopening | Saved transition attributes |
| --- | --- | --- |
| No change | true / false / 0 | none |
| Set time to 5000 | true / true / 5000 | `advTm="5000"` |
| Enable `AdvanceAfter` with default time | true / true / 0 | `advTm="0"` |
| Set time to 5000, then disable `AdvanceAfter` | true / false / 0 | none |
| Disable click advance | false / false / 0 | `advClick="0"` |
| Set time to 2147483647 | true / true / 2147483647 | `advTm="2147483647"` |

To reproduce each row, replace the settings in the iteration 1 probe with the
listed operations on `presentation.Slides[0].SlideShowTransition` and print
those three properties before saving and after reopening. Read `p:transition`
in the saved slide XML to inspect the attributes.

Aspose retains a disabled timer's previous value in memory but loses that
value after save/reopen. ShapeCrawler's nullable API clears the value
immediately; it does not retain a separate inactive timer.

The reference accepted 2147483648, but Open XML validation rejects that value:
the SDK's transition-time constraint is at most 2147483647. ShapeCrawler
therefore throws `ArgumentOutOfRangeException` before modifying XML for larger
values. The reference also treated `uint.MaxValue` as no timer when saving;
that sentinel is represented explicitly by null in ShapeCrawler.
The [Open XML transition documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition)
describes the independent `advClick` and `advTm` attributes and their defaults.

Regression tests in `SlideTransitionTests` cover defaults without XML mutation,
zero and maximum valid times, disabling, independent flags, saved XML,
save/reopen, existing effects and animations, schema ordering, both markup
compatibility branches, and out-of-range inputs. `UpdatedSlideCollectionTests`
covers importing the settings and editing the imported slide independently.
Saved presentations are Open XML validated, and `PublicApiTests` continues
to enforce SDK encapsulation. Effects, speed, sound, and other transition
options remain outside this iteration's API.

## Iteration 3: standard transition effects and speed

The narrow comparison found no ShapeCrawler equivalents for
`ISlideShowTransition.Type` and `Speed`. They are now available as
`IUserSlide.Transition.Type` and `Speed`, using ShapeCrawler-owned enums:

```csharp
slide.Transition.Type = SlideTransitionType.Fade;
slide.Transition.Speed = SlideTransitionSpeed.Slow;
slide.Transition.Type = SlideTransitionType.None; // Remove the effect.
```

The supported standard effects are Blinds, Checker, Circle, Comb, Cover, Cut,
Diamond, Dissolve, Fade, Newsflash, Plus, Pull, Push, Random, RandomBar, Split,
Strips, Wedge, Wheel, Wipe, and Zoom. The speeds are Fast, Medium, and Slow.
The defaults are None and Fast. Reading absent settings creates no XML.

Changing the type preserves click/timer settings, sound actions, extension
lists, and animation XML. Assigning the same type preserves existing effect
parameters, including direction and wheel spokes. Assigning a different type
creates it with its **schema defaults**, which can differ from Aspose's chosen
defaults (for example, direction or wheel spokes). Direction and other
effect-specific parameter APIs are deferred.

Newer effects outside the standard set return `SlideTransitionType.Unsupported`
and remain intact when changing speed or advance settings. This read-only
sentinel and invalid enum values cannot be assigned. An explicitly selected
supported type, including None, replaces an existing unsupported effect.
Type and speed updates apply to both Choice and Fallback transition branches.

### Reference evidence and differences

The reference assembly and hash remain the same as iteration 1. REA 6.7.0 was
invoked for the enum declarations:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.SlideShow.TransitionType --type Aspose.Slides.SlideShow.TransitionSpeed --json
```

Evidence `ev_83155d41387bd43dddb939bc1380e8e37c4096348d1a0bfc5f61a8264dbfcd07`
again reports unavailable member coverage at the 64 MiB metadata budget.
No implementation was recovered from that result. Public API documentation
and an unmodified Aspose evaluation-mode probe supplied the observable behavior.

The probe iterated over the 21 effects and None, set Slow plus a 5000 ms timer,
saved each presentation, inspected **all descendant** `p:transition` elements,
and reopened it with Aspose. All 22 type identities, speeds, and timers survived.
Separate Fast/Medium/Slow save/reopen cases all preserved speed. The initial
state was None/Fast. Effect element names matched the standard names, including
`randomBar` and `newsflash`.

Aspose writes some standard effects inside markup-compatibility Choice/Fallback
branches, with `p14:dur` in the Choice transition. For example, Fade/Slow was
saved with `spd="slow"` and `p14:dur="700"`. The
[Microsoft transition example](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/99b95b35-568a-4652-9cba-df3d1175952f)
explains how that extension expresses an explicit duration in milliseconds.
ShapeCrawler does not reproduce Aspose's per-effect duration table: assigning
Speed clears explicit `p14:dur` so the standard speed setting applies. Changing
the effect also clears that override; retaining the same effect preserves it.
This API does not yet expose arbitrary transition duration.

An additional compatibility probe generated **66** ShapeCrawler presentations
(22 type states times 3 speeds), each with a 5000 ms timer, and loaded them in
Aspose. Every type, speed, and timer matched. To reproduce, use a temporary
console project referencing ShapeCrawler and Aspose.Slides.NET 24.12.0:

```csharp
foreach (var type in Enum.GetValues<ShapeCrawler.SlideTransitionType>()
             .Where(type => type != ShapeCrawler.SlideTransitionType.Unsupported))
foreach (var speed in Enum.GetValues<ShapeCrawler.SlideTransitionSpeed>())
{
    using var source = new ShapeCrawler.Presentation(p => p.Slide());
    source.Slides[0].Transition.Type = type;
    source.Slides[0].Transition.Speed = speed;
    source.Slides[0].Transition.AdvanceAfterTime = 5000;
    using var stream = new MemoryStream();
    source.Save(stream);
    using var reference = new Aspose.Slides.Presentation(stream);
    var result = reference.Slides[0].SlideShowTransition;
    if (result.Type.ToString() != type.ToString() ||
        result.Speed.ToString() != speed.ToString() || result.AdvanceAfterTime != 5000)
        throw new Exception("Transition compatibility mismatch");
}
```

This verifies serialized type/speed identity and timing interoperability, not
pixel-identical visual playback. No slide images were generated. No Aspose
dependency or proprietary implementation was added to ShapeCrawler.
Regression coverage extends `SlideTransitionTests` and
`UpdatedSlideCollectionTests`, with saved-PPTX inspection, save/reopen,
schema validation, import independence, and `PublicApiTests` encapsulation.

## Candidate areas for later inventory

These areas still require member-by-member comparison against the current
ShapeCrawler API and an explicitly recorded Aspose version. They are neither
confirmed gaps nor a complete inventory:

- Slide transitions, advance timing, animations, and media playback options.
- Text formatting, lists, hyperlinks, fields, and text search/replacement.
- Shape geometry, fills, effects, grouping, and connectors.
- Chart data, axes, labels, legends, and formatting.
- Tables, cell formatting, merges, and sizing.
- Notes, comments, sections, masters, layouts, and presentation properties.
- Import/export formats and document protection, excluding slide images.

Before selecting the next implementation, inventory a narrow area, identify
existing equivalents, choose a testable missing capability, and update this
ledger in that PR. Complete coverage remains unassessed.

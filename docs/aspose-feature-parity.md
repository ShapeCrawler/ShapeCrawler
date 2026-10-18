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

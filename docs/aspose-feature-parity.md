# Aspose.Slides feature parity

This ledger tracks incremental, independently implemented ShapeCrawler features
investigated with [REA](https://github.com/morluto/rea). It records a bounded
comparison, not a claim of complete Aspose.Slides compatibility.

## Iteration rules

- Keep each PR small enough to create within 40 minutes, including investigation
  and local validation. Check all PR checks after publishing.
- Stop after each PR and wait for the maintainer's approval before the next one.
- Exclude slide-to-image rendering; it is being developed separately.
- Prioritize common editing and template workflows over additional specialized features.
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

## Iteration 4: orientation, wheel spokes, and black-screen options

The narrow comparison identified missing equivalents for
`IOrientationTransition.Direction`, `ISplitTransition.Orientation`,
`IWheelTransition.Spokes`, and `IOptionalBlackTransition.FromBlack`.
They are now exposed through the existing `ISlideTransition`:

```csharp
slide.Transition.Type = SlideTransitionType.Split;
slide.Transition.Orientation = SlideTransitionOrientation.Vertical;
slide.Transition.Type = SlideTransitionType.Wheel;
slide.Transition.WheelSpokes = 8;
slide.Transition.Type = SlideTransitionType.Fade;
slide.Transition.FromBlack = true;
```

| ShapeCrawler property | Applicable effects | Saved attribute | Missing-attribute default |
| --- | --- | --- | --- |
| `Orientation` | Blinds, Checker, Comb, RandomBar | `dir="horz"` or `dir="vert"` | Horizontal |
| `Orientation` | Split | `orient="horz"` or `orient="vert"` | Horizontal |
| `WheelSpokes` | Wheel | `spokes` (unsigned integer) | 4 |
| `FromBlack` | Fade, Cut | `thruBlk` (boolean) | false |

Getters do not create attributes. These settings leave type, speed, explicit
custom duration, advance timing, sounds, and other effect parameters intact.
All Choice/Fallback branches must support the requested parameter; otherwise
getters and setters throw `InvalidOperationException`. Setters validate the
complete branch set before making any changes. An invalid orientation enum
throws `ArgumentOutOfRangeException` without mutation. Set the applicable Type
before accessing an effect-specific property.

### Reference evidence and compatibility

The reference assembly and SHA-256 remain those recorded in iteration 1.
REA 6.7.0 was invoked with these declaring interfaces:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.SlideShow.IWheelTransition --type Aspose.Slides.SlideShow.IOptionalBlackTransition --type Aspose.Slides.SlideShow.IOrientationTransition --type Aspose.Slides.SlideShow.ISplitTransition --json
```

Evidence `ev_92e1ab4ba3729522d417c055e8af433156be5b01fdda94ce7a3737e2d8a840b1`
reports unavailable member coverage at the same 64 MiB metadata budget. No
method-body findings are inferred. Package API documentation identified the
properties; an unmodified evaluation-mode runtime probe saved each parameter
value, inspected all descendant `p:transition` elements, and reopened with Aspose.

Observed reference cases:

- Horizontal and Vertical survived save/reopen for each of the five oriented
  effects. Split used `orient`; the other four used `dir`. Aspose omitted the
  orientation attribute for Horizontal and wrote `vert` for Vertical.
- Both boolean states survived for Fade and Cut. True wrote `thruBlk="1"`;
  false omitted the attribute.
- Wheel values 0, 1, 2, 3, 4, 5, and 8 survived. Four was serialized without
  a `spokes` attribute and reopened as 4. Aspose accepted 5 as well as the
  commonly used spoke counts.

Fresh Aspose Blinds/RandomBar/Split objects initially chose Vertical, and
fresh Wheel chose 0. ShapeCrawler continues to use the schema defaults
documented in iteration 3: Horizontal and four spokes. Explicit values are
interoperable; the two libraries' fresh-effect defaults are not claimed to be
identical. The [Microsoft wheel documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.wheeltransition)
defines `spokes` as an unsigned integer, with examples for 1, 2, 3, 4, and 8.

A separate compatibility probe generated ShapeCrawler files and loaded them
with Aspose, checking the corresponding `SlideShowTransition.Value` interface
and the retained 5000 ms advance timer. All **22 cases** passed: ten orientation
cases (five effects times two orientations), four black-screen cases (two
effects times two booleans), and eight spoke counts (0, 1, 2, 3, 4, 5, 8,
and `uint.MaxValue`). Use the iteration 3 compatibility loop with the inputs
above; inspect `IOrientationTransition.Direction` or `ISplitTransition.Orientation`,
`IOptionalBlackTransition.FromBlack`, and `IWheelTransition.Spokes` respectively.

Regression coverage in `SlideTransitionTests` checks saved XML, defaults without
mutation, save/reopen, schema validity, parameter applicability, invalid enums,
compatible Choice/Fallback updates, rejection before partial branch updates,
and preservation of duration and sound. `UpdatedSlideCollectionTests` verifies
imported wheel options and source independence. Rendering remains excluded;
no images, playback comparisons, or production Aspose dependency are involved.

## Iteration 5: template text replacement

The maintainer requested prioritizing frequently used library workflows.
This iteration therefore moves from specialized transition settings to
presentation-wide template substitution. `IParagraph.ReplaceText` already
exists; the missing capabilities were text-box and presentation-level
replacement with run-spanning matching, counts, and optional case folding.
The existing paragraph API's behavior is unchanged.

```csharp
int count = presentation.ReplaceText("{{customer}}", "Alice");
textBox.ReplaceText("{{city}}", "Berlin", ignoreCase: true);
tableCell.TextBox.ReplaceText("{{address}}", "Main Street 1\nBerlin");
```

`IPresentation.ReplaceText` searches slide text boxes, nested groups, and table
cells using the existing `GetTexts` traversal. `ITextBox.ReplaceText` also
works on an explicitly accessed notes text box. The bulk operation excludes
notes, masters/layouts, chart data, and SmartArt data parts. Regex matching,
whole-word matching, and result callbacks are deferred.

The contract is literal, ordinal, non-overlapping matching against the original
text. Case-sensitive matching is the default; `ignoreCase: true` uses
OrdinalIgnoreCase. Matching spans consecutive text runs but stops at paragraph
boundaries, soft breaks, fields, or other non-run content. Each replacement
inherits the first matched run's properties. Prefix and suffix formatting,
paragraph properties, hyperlinks, font sizes, and shape geometry are retained.
The operation does not refit text or resize shapes/rows. Empty replacement
deletes matches. Replacement text containing the search text is inserted once,
without recursively replacing newly inserted content.

Replacement CRLF, LF, and CR line endings become soft breaks inside the same
paragraph. Null arguments, an empty search string, or replacement characters
that are invalid in XML are rejected before mutation. The returned integer
counts replacements, rather than text boxes changed.

### Reference evidence and limits

The reference version and assembly hash are unchanged from iteration 1.
REA 6.7.0 was invoked for `Aspose.Slides.TextFrame` and
`Aspose.Slides.TextSearchOptions`:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.TextFrame --type Aspose.Slides.TextSearchOptions --json
```

Evidence `ev_df5ea8bfc8cc2a99c227db97a5f228709e7e3ff5724031a9e83cda29ae08d5a4`
reports unavailable member coverage at the 64 MiB metadata budget. No proprietary
method bodies were recovered. The package's API documentation identifies
`TextFrame.ReplaceText` and `Presentation.ReplaceText` with search options.

An unmodified Aspose evaluation-mode probe used two portions: bold `aX` and
italic `Yz`, and searched for `XY` with CaseSensitive explicitly enabled.
The short inputs avoid evaluation-mode truncation observed with longer inputs.

| Replacement | Reference text after replacement | Formatting observation |
| --- | --- | --- |
| `Q` | `aQz` | Prefix/replacement bold; suffix italic |
| empty | `az` | Prefix bold; suffix italic |
| `XY!` | `aXY!z` | Original match replaced once |
| `Q\nR` | `aQ\nRz` in one paragraph | Replacement bold; suffix italic |

The default Aspose `TextSearchOptions.CaseSensitive` was false. ShapeCrawler
deliberately follows the case-sensitive default of ordinary .NET literal
replacement; callers opt into ignoring case. Aspose can split the prefix and
replacement into separate portions; ShapeCrawler can retain them in the same
run when the formatting is identical. Exact XML/portion identity is not claimed.

A separate interoperability probe saved the four ShapeCrawler results above,
opened them in Aspose, and checked text plus bold prefix and italic suffix.
All four passed, including the soft-break representation of the multiline value.
No evaluation restrictions were bypassed. No slide images were generated or
rendering implementation changed, and Aspose is not a project dependency.

Regression tests are in `TextBoxTests`, `TableCellTextBoxTests`, and
`PresentationTests`, with SDK inspection of saved files, schema validation,
save/reopen, mixed formatting, original-match counting, empty and multiline
values, ordinal case matching under Turkish culture, field/break boundaries,
invalid arguments, groups, tables, and multiple slides. `PublicApiTests`
continues to enforce SDK encapsulation.

## Iteration 6: whole-shape click hyperlinks

Clickable buttons, pictures, and slide navigation are common editing workflows.
ShapeCrawler already exposed hyperlinks on text portions but lacked a click-link
editor for the entire shape. The new `IShape.Link` reuses `IHyperlink`:

```csharp
shape.Link.AddFile("https://example.com/report");
pictureShape.Link.AddSlideNumber(2); // One-based index in the same presentation.
shape.Link.Remove();
```

This counterpart to Aspose's shape `HyperlinkClick`/`HyperlinkManager` works on
text shapes, pictures, tables, and groups through their nonvisual properties.
`IHyperlink.Remove` also works on existing text-portion links. Whole-shape links
and text links are independent. Reading the editor and its absent-link getters
does not create a hyperlink element or relationship.

External/file addresses use hyperlink relationships. Internal navigation uses
a slide-part relationship with `ppaction://hlinksldjump`, supports self-links,
and reports the target's current number after reordering. `File` returns null
for an internal link, and `SlideNumber` returns null for an external link.
Invalid slide numbers, empty addresses, and malformed URIs are rejected before
changing the current link. Internal-link creation is scoped to regular slides.

Replacing/removing a click link reclaims unused hyperlink relationships but
retains relationships still used elsewhere in that part, including by text or
mouse-over links. URL fragments remain distinct targets (`#one` is not `#two`).
Target slide parts are never deleted by unlinking. Repeated
removal is harmless. Text and formatting remain intact. Picture copying now
copies relationships from the cloned nonvisual properties too; the existing
image-byte copy behavior is preserved.

### Reference evidence and compatibility

The reference version and assembly SHA-256 remain those recorded in iteration 1.
REA 6.7.0 was invoked for the reference managers:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.HyperlinkManager --type Aspose.Slides.Hyperlink --json
```

Evidence `ev_51f18dc3c8b85a8ba4d3927e38a48e96abb7fad3bf314181445af9b791dd0f77`
again reports unavailable member coverage at the 64 MiB metadata budget. No
method bodies were recovered. Public API documentation and an unmodified
evaluation-mode Aspose probe supplied behavioral evidence:

| Reference operation on a rectangle | Saved shape XML | After save/reopen |
| --- | --- | --- |
| `HyperlinkManager.SetExternalHyperlinkClick(url)` | `p:cNvPr/a:hlinkClick` with relationship ID and no action | `HyperlinkClick.ExternalUrl` matched the URL |
| `HyperlinkManager.SetInternalHyperlinkClick(secondSlide)` | Same element with `action="ppaction://hlinksldjump"` | `HyperlinkClick.TargetSlide` was the second slide |
| `HyperlinkManager.RemoveHyperlinkClick()` after adding a URL | No click hyperlink element | `HyperlinkClick` was null |

To reproduce, create an Aspose presentation with two slides and a rectangle on
the first, run each manager operation, save as PPTX, inspect the first slide's
`p:cNvPr/a:hlinkClick`, and reopen to read `HyperlinkClick`. The probe used
`https://example.com/report` and left evaluation restrictions intact.

A separate probe created six ShapeCrawler presentations: text shape and picture,
each with an external URL, an internal link to slide 2, or a removed click link.
Aspose opened every file and reported the expected ExternalUrl, TargetSlide,
or null respectively (**6/6 passed**). These verify hyperlink serialization and
target resolution, not mouse interaction in PowerPoint. Hover-link authoring,
tooltips, and other action kinds are outside this increment.

Regression tests in `ShapeTests`, `HyperlinkTests`, and
`UserSlideShapeCollectionTests` inspect saved XML/relationships, reopen files,
and Open XML validate them. They cover applicable shape kinds, external/internal
and self-links, reordering, shared relationships, mouse-over preservation,
invalid input, replacement in both directions, target-slide retention, and
independent text/picture copies. Existing text-link checks and `PublicApiTests`
also pass. Rendering code is unchanged and no Aspose dependency was added.

## Iteration 7: editable shape rotation and visibility

Basic template editing and slide composition now support setting `IShape.Rotation`
and `IShape.Hidden`. Existing read APIs remain available. Rotation is normalized
to [0, 360) and rounded to the nearest Open XML angular unit (1/60000 degree).
Coordinates, sizes, and flip flags are preserved. For placeholders without a
local transform, the inherited transform is cloned locally before changing the
angle; shared layouts and masters remain unchanged. Visibility uses the shape's
nonvisual hidden flag. Rendering code is unchanged.

The reference version and SHA-256 remain those recorded in iteration 1. REA
6.7.0 was invoked with `inspect-managed-members $assembly --type Aspose.Slides.Shape --json`.
Evidence `ev_da10cfe512f2af123fbd66f783c701497bc1c0f342d95d374d908487c89d4d7a`
reports unavailable coverage at the 64 MiB metadata budget. No method bodies
were recovered. Public reference XML documentation and an unmodified Aspose
evaluation-mode probe supplied the following behavioral evidence:

| Assigned rotation | Reference rotation after save/reopen | Stored `rot` |
| --- | --- | --- |
| -450 | 270 | 16200000 |
| -45 | 315 | 18900000 |
| 0, 360, 720 | 0 | omitted |
| 405 | 45 | 2700000 |
| 33.333332 (float) | 33.333332 | 2000000 |

For each angle the probe set `Hidden` first to true and then to false or true,
saved a rectangle, inspected `a:xfrm/@rot` and `p:cNvPr/@hidden`, and reopened
the file. Hidden true was stored as 1; false was omitted. Both values survived
reopening. Aspose converts NaN/infinity to zero; ShapeCrawler deliberately
rejects nonfinite angles with `ArgumentOutOfRangeException` before mutation.

A second probe created text, picture, table, group, and line shapes using
ShapeCrawler, assigned -45 degrees and each visibility value, and opened the
saved files with Aspose. All **10/10** cases read back as 315 degrees with the
expected hidden flag. This verifies serialized metadata, not visual rendering.

Regression tests cover those shape kinds, normalization, precision, invalid
angles, inherited placeholder geometry, flip preservation, and independent
text/picture imports. Tests inspect saved packages, reopen them, and validate
Open XML. The public SDK encapsulation guard remains intact. Aspose is not a
production dependency.

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

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

## Iteration 8: presentation-wide font family replacement

### Demand research before gap selection

The research sample prioritized recurring automation workflows, excluding
rendering. It is evidence of demand, not a usage census or a claim that any
feature is universally the most popular. Existing equivalents were checked
before selecting the gap:

| Candidate | Demand evidence | Current ShapeCrawler equivalent / decision |
| --- | --- | --- |
| Extract text for processing/search | [User request for whole-presentation extraction](https://forum.aspose.com/t/extract-text-from-ppt-and-pptx/97463) and [later extraction questions](https://forum.aspose.com/t/challenges-in-extracting-text-from-slides-in-java-dealing-with-split-lines/283445) | `IUserSlide.GetTexts()` exposes slide/group/table text and can be iterated across `Slides`; skip this existing workflow. |
| Replace images in a template | [User request to replace one picture independently](https://forum.aspose.com/t/aspose-slides-for-net-replaces-all-images-in-a-presentation-instead-of-just-one/168484) | `IImage.Update(Stream)` exists; skip. |
| Replace a font family throughout a deck | [Explicit request to standardize all text fonts](https://forum.aspose.com/t/how-to-change-the-font-of-all-the-text-in-the-ppt-to-another-desired-font/166308), [another font replacement use case](https://forum.aspose.com/t/incorrect-font-rendering-and-boxed-characters-after-xlsx-pdf-pptx-conversion-in-java/329147), and [official .NET replacement API](https://docs.aspose.com/slides/net/font-replacement/) | Individual portion/theme setters exist, but no global replacement API covers all related parts. Selected gap. |

### API and behavior

`IPresentation.ReplaceFont(oldName, newName)` replaces complete, ordinal,
case-sensitive DrawingML typeface references. It traverses related parts once,
including themes, masters, layouts, notes, charts, SmartArt, groups, and table
text. Latin, East Asian, complex-script, symbol, bullet, and supplemental theme
font slots are included. Existing theme aliases remain aliases; changing their
theme definition updates inherited typography without materializing local fonts.
Text and other formatting are preserved. Null/blank names are rejected before
mutation; equal source/destination names and absent source families are no-ops.

The scope is serialized DrawingML references, not platform font fallback,
font-file embedding/subsetting, VML, or font references inside embedded Office
documents or pictures. Embedded binaries and font metadata are not renamed.
No rendering implementation or Aspose production dependency was added.

### REA investigation and behavioral evidence

The reference remains Aspose.Slides.NET 24.12.0, net6.0 assembly, SHA-256
`edaaa3005d2a462548fadcb35ffc661e2b752e7ee863c9c515f459ef7e345ca2`.

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.FontsManager --json
```

Evidence `ev_6b30cdd152d070ddd4290df79651966b7a772f10204be7acabd5885a0df5845c`
reports unavailable member coverage at the 64 MiB metadata representation
budget. A REA `decompile` attempt for `Aspose.Slides.FontsManager.ReplaceFont`
reported no available provider for the managed target. A REA `capture-process`
attempt for the compiled probe reported that native Windows PTY capture is
unsupported. Consequently, REA did not recover this method's implementation.
The following are direct observations from an unmodified evaluation-mode
reference, not decompiled implementation claims:

| Source requested for replacement with Verdana | Observed saved/reopened result |
| --- | --- |
| Arial | Arial became Verdana; arial and Arial Narrow remained unchanged. |
| arial | Only the lowercase family became Verdana. |
| Absent | Existing family references remained unchanged. |

To reproduce, create rectangles with explicit Latin, East Asian, complex-script,
and symbol fonts set to `Arial`, `arial`, `Arial Narrow`, and `+mn-lt`; add an
Arial table and notes text; save, set the theme's Latin families to Arial in the
PPTX, and reopen. Call `FontsManager.ReplaceFont(new FontData(source), new
FontData("Verdana"))`, save again, inspect DrawingML typeface attributes, and
reopen. The exact matches changed in those run slots, table/notes text, and theme
Latin slots; theme aliases remained symbolic. Evaluation restrictions remained
intact and no image rendering was invoked.

The reference probe retained some Arial supplemental script mappings. This
implementation deliberately also updates exact matches in those mappings so
that the family replacement covers explicitly authored script-specific theme
fonts. Byte-identical Aspose output is not claimed.

An interoperability probe created a ShapeCrawler deck, replaced Arial with
Verdana, and opened it in Aspose: ordinary text, two grouped children, table
text, notes Latin text, and an inherited major-theme font all read correctly
(**6/6**). The first four cases also verified effective East Asian fonts; notes
were checked for their explicitly authored Latin reference, not Aspose's implicit
East Asian fallback.

Tests in `PresentationTests` cover those workflows, complete-name matching,
case sensitivity, absent/same-name families, invalid input, script/symbol/bullet
fonts, supplemental maps, inherited aliases, and saved/reopened packages.
Chart/table/group fixtures also include SmartArt and notes, validate Open XML,
and verify byte preservation of binary parts. `PublicApiTests` checks that the
new API exposes no SDK types.

## Iteration 9: custom document properties

### Demand and gap check

Custom metadata is a recurring reading/editing/creation workflow: users asked for
[an API to add custom presentation properties](https://forum.aspose.com/t/api-to-add-custom-property/86872),
[consistent custom-property CRUD across Office libraries](https://forum.aspose.com/t/why-the-api-diff-between-slides-amp-words-cells-wrt-to-customproperties/91555),
and [metadata updates in a cloud download workflow](https://forum.aspose.com/t/save-word-excel-ppt-file-without-changing-last-modified-date-user/179214).
The [official .NET properties guide](https://docs.aspose.com/slides/net/presentation-properties/)
provides custom name/value authoring and enumeration. These are demand signals,
not a complete usage ranking.

The current API already exposes built-in `IPresentationProperties` fields and
slide-level `IUserSlide.CustomData`. Neither provides document-level custom
property name/value access. A source search found no custom-property API; this
is the selected missing capability. Rendering is outside the research scope.

### API and serialization

```csharp
presentation.Properties["ReviewedBy"] = "Alice";
presentation.Properties["Version"] = 2;
var reviewer = (string)presentation.Properties["reviewedby"];
var names = presentation.Properties.CustomPropertyNames;
var exists = presentation.Properties.ContainsCustomProperty("Version");
var removed = presentation.Properties.RemoveCustomProperty("Version");
presentation.Properties.ClearCustomProperties();
```

The indexer supports string, int, bool, finite float/double, and DateTime.
Names use ordinal case-insensitive matching; original spelling survives an
update and names are returned in case-insensitive ordinal order. Empty and
whitespace names are supported, as in the reference. A missing indexer read
throws `KeyNotFoundException`; removing a missing property returns false.
Clearing is idempotent and leaves built-in properties intact.

Properties are stored in `docProps/custom.xml` with the standard format ID,
unique IDs starting at 2, and `lpwstr`, `i4`, `bool`, `r4`, `r8`, or `filetime`
variants. Existing IDs and unrelated properties are preserved when a value's
type changes. Reads do not create a part; the first successful write creates it.
Removal/clear deletes an empty custom-property part, and adding again works
through the same properties object. Null/unsupported values and invalid XML
characters are rejected before mutation. Existing unsupported variants are
preserved during unrelated edits and can be listed/removed; reading their value
throws `NotSupportedException`. Legacy `lpstr` and `bstr` strings are also read.

Dates are stored and returned as UTC. Unspecified dates are treated as UTC;
local dates are converted to UTC. The reference probe relabeled local wall-clock
times as UTC instead, so this conversion is a deliberate difference. The
reference also renumbered property IDs when saving; this implementation preserves
existing IDs. Byte-identical reference output is not claimed.

### REA investigation and reference observations

REA 6.7.0 member inspection targeted `Aspose.Slides.DocumentProperties` in both
assemblies from Aspose.Slides.NET 24.12.0:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.DocumentProperties --json
```

| Artifact | SHA-256 | Evidence |
| --- | --- | --- |
| net6.0 (also used by runtime probes) | `edaaa3005d2a462548fadcb35ffc661e2b752e7ee863c9c515f459ef7e345ca2` | `ev_7343392944608ca17d0b4b41ae719045c743f7b9811f23f52f5847ce11c939bb` |
| netstandard2.0 | `004750d1968ccc1a05a9a6985293a11c42a923e8f9b51bf46ff545cded7a541e` | `ev_d48fa8e9cfa7f1e307aedc386bef671834c299934a2b52ea7e6b6bc53f6c8585` |

Both report unavailable member coverage at the 64 MiB metadata representation
budget. REA recovered no method bodies; behavioral observations below come from
direct calls to the unmodified evaluation-mode reference and inspection of its
saved PPTX, supplemented by the reference package's XML API documentation.

To reproduce, create an Aspose presentation and set document properties to a
string containing `&` and `<`, positive/negative integers, true/false, 1.25f,
1.23456789012345d, and a UTC DateTime. Read their runtime types, save, inspect
`docProps/custom.xml`, reopen, and read again. The six types survived with the
variants listed above. Int64, decimal, and null assignments threw
`PptxEditException`. Write `Key`, `key`, and `KEY`: they update one property while
retaining the original name. Remove with `KEY`, then repeat removal of a missing
name: no exception occurs. `ClearCustomProperties()` yields a count of zero.
Empty/whitespace names survive save/reopen. Evaluation restrictions were retained
and no slide image generation was invoked.

An interoperability probe verified the six supported types in both directions:
ShapeCrawler-created files opened in Aspose, and Aspose-created files opened in
ShapeCrawler (**12/12 typed checks passed**). Cross-library update, case-insensitive
removal, clear, and built-in title preservation also passed.

`PresentationPropertiesTests` covers typed save/reopen under de-DE culture,
XML escaping, names/case matching, type changes, unique/stable IDs, missing
properties, idempotent removal/clear, UTC dates, rejection without mutation,
unsupported-value preservation, and Open XML validation. `PublicApiTests`
continues to enforce SDK encapsulation. No Aspose production dependency or
rendering changes are included.

## Iteration 10: literal and regex text highlighting

### Demand and gap check

A user explicitly requested [regex highlighting across a complete presentation](https://forum.aspose.com/t/highlight-and-list-all-the-highlighted-words-in-powerpoint-presentations-in-c/280468),
including counts/reporting. Another asked about [editing text highlight colors](https://forum.aspose.com/t/working-with-highlight-effect-for-text-portions-in-powerpoint-presentations-in-java/264416).
The [official .NET search guide](https://docs.aspose.com/slides/net/search-and-replace-text/)
documents frame-level and presentation-wide literal/regex highlighting. These
support prioritizing a recurring review/editing workflow; no universal usage
ranking is claimed.

The existing `IParagraphPortion.TextHighlightColor` edits an entire portion.
`ReplaceText` changes content and `GetTexts` collects text boxes. None finds
matches and highlights only those characters, including partial/cross-run
matches, while retaining their original content and formatting. A source search
confirmed the absence of `HighlightText`/`HighlightRegex` APIs before this change.
This is the selected missing capability, not a rendering feature.

### API and scope

`ITextBox` and `IPresentation` now expose:

```csharp
int count = presentation.HighlightText("Confidential", new Color("FFFF00"), ignoreCase: true);
int emails = presentation.HighlightRegex(emailRegex, new Color("FFFF00"));
```

Both return non-overlapping match counts. Literal matching is ordinal and
case-sensitive by default, consistent with ShapeCrawler's `ReplaceText`; set
`ignoreCase: true` for case-insensitive matching. Regex matching follows the
supplied .NET Regex options and timeout, supports RightToLeft, and ignores
zero-length matches. Adjacent match spans share highlighting segments while
still counting as separate matches.

Matches can span consecutive ordinary runs. Paragraphs, soft breaks, fields,
and other non-run nodes are boundaries; fields are not searched. Presentation
scope covers slide text boxes, nested groups, and table cells through `GetTexts`.
Notes, masters, layouts, charts, and SmartArt data are excluded from that scope;
individual accessible text boxes (including notes) can be processed explicitly.
Callbacks, match-result reporting, and dedicated whole-word options are outside
this increment; word boundaries can be expressed in the regex.

Runs are split only at selected boundaries; text, font/style properties,
hyperlinks, paragraph properties, and outside highlights remain intact. Existing
relationship IDs are reused within the same part. Highlight insertion respects
schema order, including runs containing hyperlink properties. Finite alpha in
[0,255] is required. `Color.NoColor` removes direct highlighting from matched
characters; it does not override an inherited paragraph highlight. Matches that
would split a Unicode surrogate pair are rejected; adjacent regex matches are
merged before this check. Matching completes before changing a text box, so a
regex timeout leaves that box unchanged. A later failure in presentation scope
can leave earlier boxes updated, as stated in the API documentation.

### REA and behavioral evidence

Reference: Aspose.Slides.NET 24.12.0 net6.0, SHA-256
`edaaa3005d2a462548fadcb35ffc661e2b752e7ee863c9c515f459ef7e345ca2`.
REA 6.7.0 was invoked as follows:

```powershell
npx --yes rea-agents@6.7.0 inspect-managed-members $assembly --type Aspose.Slides.TextFrame --type Aspose.Slides.TextSearchOptions --json
```

Evidence `ev_df5ea8bfc8cc2a99c227db97a5f228709e7e3ff5724031a9e83cda29ae08d5a4`
reports unavailable member coverage at the 64 MiB metadata representation
budget. No method implementation was recovered. The following observations are
from direct calls to the unmodified evaluation-mode reference and inspection of
saved XML, with API signatures checked against the package's XML documentation:

Create a text frame with a bold 18-point `a XY` portion and italic
`Z xyz X12Z tail` portion. Highlight `XYZ` with `CaseSensitive=true`, save, and
inspect the paragraph: `a `, `XY`, `Z`, and ` xyz X12Z tail` are separate runs,
with highlight only on XY/Z and the original font properties preserved.
Default reference highlighting of `xyz` also selects uppercase XYZ, demonstrating
its case-insensitive default. Regex `X[0-9]+Z` selects only X12Z. Yellow is stored
as `a:highlight/a:srgbClr` with FFFF00; alpha 128 is stored as `a:alpha val=50196`.
Reopening reports the expected colors on matched portions. Reference evaluation
getters truncate long portions; full original text was therefore checked in the
saved XML, without bypassing evaluation restrictions.

Known differences: ShapeCrawler defaults to case-sensitive literal matching and
returns a count rather than using a callback. Reference transparent highlighting
writes an explicit alpha-zero highlight; ShapeCrawler uses the existing ShapeCrawler
NoColor convention of removing the direct highlight. No byte-identical output or
complete search/highlighting API parity is claimed.

An interoperability probe created ShapeCrawler text, grouped text, and table
cells, each with literal, regex, alpha-128, or cleared direct highlighting.
Aspose opened all **12/12** cases, preserving the complete short text and
reporting the expected highlight RGBA on matched portions. No slide image was
generated.

Tests in `TextBoxTests`, `TableCellTextBoxTests`, and `PresentationTests` cover
partial/cross-run matches, original non-overlapping counts, case matching under
tr-TR, RightToLeft/empty regex matches, Unicode, field/break boundaries, alpha,
repeated calls, direct-highlight removal, hyperlink/format preservation, timeout
without partial text-box mutation, excluded notes/masters, saved/reopened output,
and Open XML validation. `PublicApiTests` checks SDK encapsulation. No Aspose
production dependency or renderer implementation changes were added.

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

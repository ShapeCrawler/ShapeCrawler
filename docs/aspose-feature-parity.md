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

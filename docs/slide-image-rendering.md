# Slide image rendering investigation

## Scope and delivery

Build out ShapeCrawler's existing experimental `SaveImageTo` renderer using REA static evidence and direct behavioral comparisons with Aspose.Slides. Each implementation PR must be small enough to investigate, implement, validate and create within 30 minutes. After each PR, wait for the user's explicit approval before starting the next.

The first increment supports explicit PNG pixel dimensions through stream and file overloads. It scales the canvas before drawing slide shapes, so shapes and text render at the requested resolution. Backgrounds cover the output surface. The existing overloads keep their 96 DPI dimensions and original canvas transform. No Aspose dependency or proprietary implementation code is added to ShapeCrawler.

## Target and evidence (2026-10-18)

The reference is the locally available NuGet package `Aspose.Slides.NET` 24.12.0, `lib/net6.0/Aspose.Slides.dll`. This is a version-specific investigation, not a claim about the latest Aspose release.

- Assembly identity: `Aspose.Slides`, version `24.12.0.0`.
- SHA-256: `edaaa3005d2a462548fadcb35ffc661e2b752e7ee863c9c515f459ef7e345ca2`.
- REA version: `rea-agents@6.7.0`, managed static provider version 1.
- Artifact inspection: `ev_6d927c317492f227a3faaa17da3ed8a04d096e61665af76e7b970a0b6cd4c818`, complete coverage.
- Stock selected-member inspection: `ev_1e0a865f6c3dfa282f28e4bac8c746817303a976617dc5bc099e99414651feaa`, unavailable because the metadata representation exceeded 64 MiB.
- Successful selected-member inspection: `ev_68679ee9f379cb7342237cb4f3e42bb36a318deb61990ed78e96a580953fe7d6`, complete coverage of the 71 declared methods in `Aspose.Slides.Slide`. This used an isolated temporary copy of REA with only `MAX_METADATA_REPRESENTATION_BYTES` increased from 64 to 256 MiB in its managed decoder (including the bundled CLI copy). The installed package and client configuration were not changed.

REA commands, with an absolute assembly path substituted for `<assembly>`:

```powershell
npm exec --yes --package=rea-agents@6.7.0 -- rea inspect-managed-artifact <assembly> --json
npm exec --yes --package=rea-agents@6.7.0 -- rea inspect-managed-members <assembly> --type Aspose.Slides.Slide --json
```

For the successful inspection, invoke `node <temporary-copy>/scripts/rea.mjs inspect-managed-members ...` after increasing the decoder budget. Save complete JSON outside the repository. REA evidence IDs identify those records; they are not remote download URLs.

### Observed static structure

REA decoded the following CIL calls. Tokens are local to the assembly digest above.

| Entry | Token | Observed call |
| --- | --- | --- |
| `Slide.GetImage(float, float)` | `0x060200e8` | `GetThumbnail(float, float)`, `0x060200e7` |
| `Slide.GetImage()` | `0x060200ea` | `GetThumbnail()`, `0x060200e9` |
| `Slide.GetImage(Size)` | `0x060200ed` | `GetThumbnail(Size)`, `0x060200ec` |
| `Slide.GetImage(IRenderingOptions, Size)` | `0x060200f4` | `GetThumbnail(IRenderingOptions, Size)`, `0x060200f3` |

Each listed image entry also constructs an image wrapper via `0x06008a49`. This establishes delegation, not the complete rasterization algorithm. Many implementation names are obfuscated, and the selected-type inspection does not decode bodies of downstream types.

### Observed runtime behavior

A separate .NET probe executed the actual Aspose assembly on Windows. REA's process-capture workflow does not support native Windows, so the probe ran directly through `dotnet`; its observations are separate from the static REA records.

For a slide of 120 by 90 points:

| Aspose request | Observed image pixels |
| --- | --- |
| `GetImage(1f, 1f)` | 120 by 90 |
| `GetImage(2f, 2f)` | 240 by 180 |
| `GetImage(0.5f, 0.5f)` | 60 by 45 |
| `GetImage(2f, 0.5f)` | 240 by 45 |
| `GetImage(0.333f, 0.777f)` | 40 by 70 |
| `GetImage(new Size(240, 45))` | 240 by 45 |
| `GetImage(new Size(45, 240))` | 45 by 240 |
| `GetImage(new Size(41, 37))` | 41 by 37 |

The comparison fixture is created with ShapeCrawler: 120 by 90 points, red slide background, blue rectangle at `(30, 20)` with size `(60, 40)`. Save it as PPTX and load that same file into Aspose. For each requested size, save the Aspose `GetImage(Size)` result as PNG and compare it with ShapeCrawler `SaveImageTo(file, width, height)`.

Both engines produced the requested dimensions and red pixels at normalized positions `(1/8, 1/2)`, `(7/8, 1/2)`, `(1/2, 1/10)`, `(1/2, 9/10)`. At 240 by 45 and 120 by 90 both produced a blue center pixel. Aspose's evaluation watermark visibly covers the rectangle and changes the center pixels in the other two outputs; those pixels do not establish a renderer mismatch or an equivalence claim. No watermark removal or license bypass was performed.

ShapeCrawler's automated regression tests reproduce the dimensions and sample positions independently of Aspose, check picture-background coverage at smaller and larger sizes, and check invalid dimensions and file preservation. Positive sizes are an explicit ShapeCrawler API contract; invalid-size behavior has not been compared with Aspose.

The different default pixel densities are intentional compatibility behavior: Aspose's observed scale-one output uses one pixel per point; ShapeCrawler's existing default uses 96 pixels per inch. Explicit pixel dimensions provide a common comparison surface.

## Remaining increments

Select one bounded behavior per approved PR. Investigate it with a small fixture before implementation, retain artifact identities and relevant REA calls, and add saved-output or pixel regression tests. Split a topic further when needed to stay inside the time limit.

1. Inherited slide backgrounds and theme-color resolution.
2. Master/layout drawing order and placeholder visibility.
3. Additional preset geometry, then custom paths.
4. Fill and outline variants, including gradients and transparency.
5. Text layout, font selection and shaping, one discrepancy at a time.
6. Nested transforms, picture cropping and effects.
7. Tables, charts and remaining graphic objects, split by feature.

These are open investigation areas, not established equivalence. Matching a few pixel samples does not prove pixel-identical rendering of arbitrary slides. The full feature remains incomplete until the agreed rendering coverage has been implemented and verified.

Sources: [REA](https://github.com/morluto/rea), [managed-code inspection](https://github.com/morluto/rea/blob/main/docs/managed-code-analysis.md), [process capture host support](https://github.com/morluto/rea/blob/main/docs/process-capture.md), and the XML API documentation shipped with the reference NuGet package.

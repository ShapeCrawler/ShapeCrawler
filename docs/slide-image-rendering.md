# Slide image rendering investigation

## Scope and delivery

Build out ShapeCrawler's existing experimental `SaveImage` renderer using REA static evidence and direct behavioral comparisons with Aspose.Slides. Each implementation PR must be small enough to investigate, implement, validate and create within 30 minutes. After creating a PR, wait for every check on its current head to pass and fix any failures in that PR. Then wait for the user's explicit approval before starting the next increment.

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

The comparison fixture is created with ShapeCrawler: 120 by 90 points, red slide background, blue rectangle at `(30, 20)` with size `(60, 40)`. Save it as PPTX and load that same file into Aspose. For each requested size, save the Aspose `GetImage(Size)` result as PNG and compare it with ShapeCrawler `SaveImage(file, width, height)`.

Both engines produced the requested dimensions and red pixels at normalized positions `(1/8, 1/2)`, `(7/8, 1/2)`, `(1/2, 1/10)`, `(1/2, 9/10)`. At 240 by 45 and 120 by 90 both produced a blue center pixel. Aspose's evaluation watermark visibly covers the rectangle and changes the center pixels in the other two outputs; those pixels do not establish a renderer mismatch or an equivalence claim. No watermark removal or license bypass was performed.

ShapeCrawler's automated regression tests reproduce the dimensions and sample positions independently of Aspose, check picture-background coverage at smaller and larger sizes, and check invalid dimensions and file preservation. Positive sizes are an explicit ShapeCrawler API contract; invalid-size behavior has not been compared with Aspose.

The different default pixel densities are intentional compatibility behavior: Aspose's observed scale-one output uses one pixel per point; ShapeCrawler's existing default uses 96 pixels per inch. Explicit pixel dimensions provide a common comparison surface.

## Second increment: inherited backgrounds

When the slide has no background element, export now uses its layout's background, or its master's background when the layout also has none. A present background stops the search: an explicit `NoFill` renders white instead of revealing the parent's background. Direct RGB fills, direct supported scheme colors such as `accent1`, and picture backgrounds use the same rendering paths as slide-owned fills. Image relationships are resolved in the part that owns the background.

Rendering reads the existing background elements directly. It no longer calls the editable `IUserSlide.Fill` getter, which can insert slide background properties and thereby override an inherited background. The setter/getter contract of `Fill` is unchanged; this increment only makes image export read the existing background hierarchy without modifying it.

### REA evidence

The reference assembly and isolated 256 MiB REA decoder configuration are the same as in the first increment. New records:

- `ev_254e64fae2b3e0dca6679882d2e22be685caf840c5cbe0cc47197148f8282a05`: complete selected-type inspection of the 23 declared methods of `Aspose.Slides.Background`.
- `Background.GetEffective`, token `0x0601da92`, delegates to `0x0601da9a`. That method caches an object constructed by `0x060019a4` and invokes the effective-data update through member reference `0x0a003bde`.
- `ev_074f1609013728212508c337d013f2c470402d541d7654a1a0ad90ddd42282e1`: complete inspection of the 13 declared methods of that effective-data object's type. Its declaring type name is obfuscated: `"\u0002\u0010\u001b\u001a"` (JSON escaping). Select the exact `target_declaring_type` reported for constructor `0x060019a4` rather than guessing a readable class name.
- Effective-data method `0x060019ad` calls `LayoutSlide.get_MasterSlide` (`0x0601f709`), `BaseSlide.get_Background` (`0x0601db1b`) and `Background`'s cached effective-data method (`0x0601da9a`). Method `0x060019ae` reads the background's type, fill format, style color and style index. These are observed CIL relationships; the runtime fixtures below establish the tested precedence and output colors.

### Behavioral comparison and regression coverage

Fixtures start from a saved ShapeCrawler PPTX. In test/probe code, the Open XML SDK removes existing slide, layout and master backgrounds and inserts the scenario's background properties. Both engines open the same resulting PPTX and export at 120 by 90 pixels. Sample `(118, 88)`, outside the reference evaluation watermark:

| Slide background | Layout background | Master background | Aspose output | ShapeCrawler before | ShapeCrawler after |
| --- | --- | --- | --- | --- | --- |
| absent | absent | absent | white | white | white |
| absent | absent | blue | blue | white | blue |
| absent | green | blue | green | white | green |
| red | green | blue | red | red | red |
| `NoFill` | green | blue | white | white | white |
| absent | absent | `NoFill` | white | white | white |
| absent | absent | `accent1` | `4472C4` | white | `4472C4` |

Automated regressions also check layout `NoFill`, image backgrounds owned by a layout or master, and unchanged slide/layout/master XML after exporting. The picture fixture deliberately assigns the same image relationship ID to different pictures in the slide and the background-owning part, to detect resolution against the wrong part. Fixture image generation does not require fonts.

A direct picture comparison used a 32 by 16 pixel PNG with red and blue halves, attached to the layout or master with a different slide-owned image under the same relationship ID. Both engines select the correct image: samples `(20, 80)` and `(100, 80)` in the 120 by 90 output are red and blue. There is a separate residual difference in image-edge resampling: with a 2 by 1 source, Aspose mixes colors at the edges while the current Skia path samples solid colors; even the larger source has a pale border in the Aspose output. This increment preserves existing image sampling and does not claim pixel equivalence at image edges.

The second increment left theme background-style lookup, color-map overrides, gradients, alpha compositing and background effects open. Unsupported explicit backgrounds use the renderer's white fallback and do not inherit a different background. The next increment addresses style lookup and color maps below.

## Third increment: theme background styles and color maps

Background rendering now has a dedicated internal `SlideBackground` implementation. It resolves `p:bgRef` against the master theme's style matrix, substitutes the reference color for `phClr`, and resolves scheme colors through the effective slide/layout/master color map. Solid RGB and system fallback colors are supported. Picture styles retain their theme-part image relationships. Export still leaves the slide, layout, master and theme XML unchanged.

### Evidence and the specification decision

The reference assembly SHA-256 was rechecked and matches the first increment. The existing complete REA record `ev_074f1609013728212508c337d013f2c470402d541d7654a1a0ad90ddd42282e1` was reused rather than decoding the same type again. Method `0x060019af` (CIL SHA-256 `9079eb1ec85b225fdd0dbe47a5666848a17eeee2ca8215f90a9f811ef7b30319`) calls `IThemeable.CreateThemeEffective`, `IThemeEffectiveData.get_FormatScheme` and `IFillFormatCollectionEffectiveData.get_Item`. Its decoded CIL includes one subtraction and a constant one. This supports tracing style lookup; the original static record alone does not establish every index or color-map rule.

Direct execution revealed a difference from the documented OOXML rules for indices 1–999. The user explicitly chose OOXML behavior for this reconstruction. The [BackgroundStyleReference documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.backgroundstylereference?view=openxml-3.0.1) specifies: 0 and 1000 mean no background; 1–999 select `fillStyleLst`; 1001 and above select `bgFillStyleLst`, starting at the first entry. This implementation follows those rules, with a white fallback for missing entries.

The fixture sets the ordinary fill styles to red, `phClr`, blue; background fill styles to `phClr`, green, `accent3`; and the reference color to `CC00CC`. A slide with the explicit reference is rendered at 120 by 90 pixels; sample `(100, 80)` lies outside the evaluation watermark:

| Index | Aspose 24.12 | ShapeCrawler after | Basis |
| --- | --- | --- | --- |
| 0 or 1000 | white | white | no background |
| 1 | `CC00CC` | red | ordinary fill style 1, per OOXML |
| 2 | green | `CC00CC` | ordinary fill style 2, per OOXML |
| 3 | `778899` | blue | ordinary fill style 3, per OOXML |
| 1001 | `CC00CC` | `CC00CC` | background style 1 with placeholder substitution |
| 1002 | green | green | background style 2 |
| 1003 | `778899` | `778899` | background style 3 with theme color |

All nonwhite cases in this table rendered white before this increment. The differences in the first three positive indices are intentional and are covered by tests using different ordinary and background style lists.

### Color maps

The map fixture assigns `bg1` to `accent1` (`112233`) in the master, `accent2` (`445566`) in the layout override, and `accent3` (`778899`) in the slide override. Both engines use the layout override when no slide override is present, and the slide override when present. This also applies when the background is inherited from a layout or master and when the reference color supplies `bg1` through `phClr`.

With explicit slide `a:masterClrMapping`, Aspose still uses the layout override in the tested fixture. ShapeCrawler uses the master's `112233`, following the documented [MasterColorMapping rule](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.mastercolormapping?view=openxml-3.0.1). This is a second recorded specification/reference difference. A layout that itself selects master mapping also resolves to `112233` in both engines.

New regressions check ordinary/background style indices, no-background and missing-index behavior, inherited references, placeholder substitution, map precedence, theme-owned pictures, and XML preservation. Existing image rendering and public API checks remain in place. No public SDK objects or new public APIs are introduced.

The third increment left theme override parts, color transforms, gradients and effects open. The master theme is the current style/color source. The next increment addresses selected color transforms below; the picture edge-resampling difference remains unchanged.

## Fourth increment: background color transforms

Solid backgrounds now apply `tint`, `shade`, `lumMod`, `lumOff`, `alpha`, `alphaMod` and `alphaOff` in XML order. This covers slide-owned and inherited RGB/system/theme colors, plus the reference color and `phClr` in theme styles. Reference-color transforms run before transforms on the style's placeholder color. Internal `DrawingColor` retains floating-point channels through the chain, then quantizes once for Skia output. No public API changes are needed.

Tint and shade operate on linear sRGB channels, with the standard piecewise sRGB transfer functions used to decode and encode. Luminance changes preserve HSL hue/saturation; neutral colors remain neutral. Luminance and alpha are clamped after each applicable operation. Solid background alpha is composited over white for image export, matching the reference's PNG output rather than emitting a transparent solid background.

### Static evidence

The reference is the same Aspose.Slides 24.12.0 assembly and isolated REA 6.7.0 decoder configuration used above:

- `ev_ee25f1e64fc3f2649c0ea06185e934de320c14c79e3e2ba0a3bb16d9fbda427c`: complete selected-type inspection of 73 declared `Aspose.Slides.ColorFormat` methods. `get_ColorTransform`, token `0x0601e408`, calls the related internal method `0x06002643`.
- `ev_7c3b2e5272bdee7b7320c7a6bac4f9af3f488a118c244ee9b8b9940d10594c5a`: complete inspection of that internal type's 62 declared methods. Method `0x06002643` returns field `0x04000d5c`; the type also contains indexed `IColorOperationCollection.get_Item` and collection count calls. This locates operation storage/access but does not, by itself, prove the numerical transform algorithm.

The implementation math was derived from the color model and corroborated by direct execution. It does not copy proprietary implementation code. Documentation describes [tint](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tint?view=openxml-3.0.1), [shade](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shade?view=openxml-3.0.1), and [luminance modulation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminancemodulation?view=openxml-3.0.1); the [DrawingML primer](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf) also describes luminance/alpha transforms and their limits.

### Runtime results

The probe starts with color `4080C0`, writes the transforms through the Open XML SDK, then opens the same PPTX in both engines. It records Aspose's effective color separately from exported pixels: for example, 50% alpha has effective ARGB `804080C0`, but the PNG sample is opaque `9FBFDF` after white compositing. At `(100, 80)` in 120 by 90 output, outside the evaluation watermark, both engines now produce:

| Operations, in order | PNG RGB |
| --- | --- |
| shade 50% | `2C5C8C` |
| tint 50% | `C0CDE2` |
| luminance modulation 50% | `204060` |
| luminance offset +25% | `A0C0E0` |
| luminance modulation 80%, offset +20% | `6699CD` |
| luminance offset +20%, modulation 80% | `578FC8` |
| alpha 50% | `9FBFDF` |
| alpha 0% | `FFFFFF` |
| alpha 100% | `4080C0` |
| alpha 80%, alpha modulation 50% | `B3CCE6` |
| alpha 25%, alpha offset +25% | `9FBFDF` |
| tint 50%, shade 50% | `8C96A6` |
| shade 50%, tint 50% | `BEC4D0` |

The scheme-color luminance fixture also yields `6699CD`. A reference color with tint 50% followed by placeholder shade 50% yields `8C96A6`. All 15 runtime comparison cases match at the selected sample; before this increment all returned unmodified `4080C0` from ShapeCrawler (the alpha-100% case is an unchanged control).

Automated saved-PPTX tests cover those operations and order differences, neutral black/white/gray colors, clamping, inherited/system/scheme colors, placeholder composition, and unchanged XML containing transforms. Fixtures do not depend on fonts.

Remaining color work includes explicit hue/saturation/component/gamma operations, other input color models and theme override parts. Transforms on shape fills, picture alpha effects, gradients and background effects are separate work. Pixel equivalence beyond the tested samples is not established.

## Remaining increments

Select one bounded behavior per approved PR. Investigate it with a small fixture before implementation, retain artifact identities and relevant REA calls, and add saved-output or pixel regression tests. Split a topic further when needed to stay inside the time limit.

1. Theme override parts and remaining color operations/models.
2. Master/layout drawing order and placeholder visibility.
3. Additional preset geometry, then custom paths.
4. Fill and outline variants, including gradients and transparency.
5. Text layout, font selection and shaping, one discrepancy at a time.
6. Nested transforms, picture cropping and effects.
7. Tables, charts and remaining graphic objects, split by feature.

These are open investigation areas, not established equivalence. Matching a few pixel samples does not prove pixel-identical rendering of arbitrary slides. The full feature remains incomplete until the agreed rendering coverage has been implemented and verified.

Sources: [REA](https://github.com/morluto/rea), [managed-code inspection](https://github.com/morluto/rea/blob/main/docs/managed-code-analysis.md), [process capture host support](https://github.com/morluto/rea/blob/main/docs/process-capture.md), and the XML API documentation shipped with the reference NuGet package.

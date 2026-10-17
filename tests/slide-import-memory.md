# Slide import memory regression (#1311)

`Slides.Add(sourceSlide)` and `Slides.Add(sourceSlide, position)` import from the
source package without cloning the entire presentation first. Unsaved source
edits are flushed before copying. Layout pruning for the append overload happens
on the destination copy, so the source retains its layouts. Only the new master's
layout IDs are assigned and saved; existing masters are no longer rewritten on
every append.

The ordinary `UpdatedSlideCollectionTests` cover both overloads, unsaved edits,
source preservation, and allocations smaller than an unrelated 8 MiB media part.
The allocation test fails with the previous whole-package clone implementation.

## Reproduce constrained memory

Build first, then launch the explicit tests in a separate process with a 1 GiB
managed heap limit. From the repository root in PowerShell:

```powershell
dotnet build ShapeCrawler.Tests/ShapeCrawler.Tests.csproj
$previousLimit = $env:DOTNET_GCHeapHardLimit
try {
    $env:DOTNET_GCHeapHardLimit = '0x40000000'
    dotnet test ShapeCrawler.Tests/ShapeCrawler.Tests.csproj --no-build --no-restore --filter 'FullyQualifiedName~UpdatedSlideCollectionTests.Add_imports_more_than_2500' --logger 'console;verbosity=normal'
} finally {
    $env:DOTNET_GCHeapHardLimit = $previousLimit
}
```

Each case generates a two-slide source with 448 MiB of opaque media attached to
the second slide, then imports the first slide 2,501 times. It saves, reopens,
checks the 2,502 destination slides, and validates the saved package using the
Open XML SDK. The import writer is disposed and collected before validation loads
another XML DOM under the same heap budget. The fixture writes the payload through ZIP compression in small
blocks; it does not allocate a 448 MiB setup buffer. The payload is deliberately
compressible and is a synthetic media part, not a displayable PNG. This isolates
unnecessary package copying from image decoding and keeps the fixture small.

This is a deterministic managed-memory equivalent, not an exact simulation of
32-bit virtual address fragmentation or a 448 MiB compressed PPTX. The old
whole-package clone inflates the unrelated part and exhausts this heap budget;
the import must leave it untouched. Tests are explicit because the stress run
takes longer than the ordinary suite.

## Remaining limits

There is no fixed 2,500-slide limit. Available memory, media size, XML complexity,
and linked parts determine the practical limit. ShapeCrawler still holds source
and destination packages in memory; importing related media and saving can need
additional buffers. Repeated append imports also preserve separate imported
masters, so their memory requirements differ from positional imports. The
positional stress case also passes with a 512 MiB managed heap budget. Both
overloads are tested with a 1 GiB budget.
Removing the whole-source clone cannot guarantee that every large deck
will fit in a 32-bit process.

For large presentations, target x64, disable Prefer 32-bit in the host application,
dispose presentations when finished, and split the output into smaller decks if
memory remains insufficient. Saving to a file avoids an additional caller-owned
output `MemoryStream`, but is not a streaming or constant-memory import API.

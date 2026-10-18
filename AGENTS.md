# Project instructions

## Temporary files

- Store temporary files created by Codex, including build/test logs, diagnostic output, scratch scripts, and intermediate artifacts, under `C:\temp` (prefer a task-specific subfolder), not in the repository.
- Create the temporary directory when needed. Keep only intended project deliverables in the repository.

## Encapsulation

- Keep Open XML SDK elements, parts, and packages internal. Do not expose them through public interfaces, constructors, methods, properties, or fields, including collection and generic wrappers.
- Do not introduce public SDK escape hatches such as `SdkOpenXmlElement`, `GetSdkPresentationDocument()`, or `GetSdkPresentationPart()`. Use internal implementation members for operations that require SDK access.
- Tests that inspect XML must save the presentation through `Save(Stream)` and open the resulting PPTX with the Open XML SDK in test code. Do not add public production APIs or widen implementation visibility solely to support tests.
- Preserve the encapsulation check in `ShapeCrawler.Tests/PublicApiTests.cs` when changing the public API.

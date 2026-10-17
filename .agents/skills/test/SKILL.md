---
name: test
description: Write or modify ShapeCrawler tests using the project's test class naming and Arrange, Act, Assert conventions.
---

# Test conventions

## Test class and file names

- Name each test class after the existing production class it tests, followed by `Tests`; the file name must match the test class name.
- Before adding a test class, locate the production class and its existing test class. Add tests to the existing test class and file when present.
- For example, tests for `Presentation`, including opening files and streams, belong in `PresentationTests` in `PresentationTests.cs`. Do not create `PresentationOpeningTests` when there is no production class named `PresentationOpening`.
- Before finishing, verify that each new test class name maps to an existing production class and that its file name matches.

## Test sections

For every new or modified test in this project:

- Always add `// Act` immediately before the operation being tested.
- Always add `// Assert` immediately before checking the result. Saving and reopening a presentation to inspect the result belongs in this section.
- Use the exact comment spelling `// Act` and `// Assert`, matching the existing C# test style.
- Keep setup before the Act section; mark it with `// Arrange` for consistency with the project's tests.
- Before finishing, check that every test added or modified has both required comments and that each comment marks the correct section.

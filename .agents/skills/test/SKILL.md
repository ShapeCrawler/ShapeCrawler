---
name: test
description: Write or modify ShapeCrawler tests using the project's required Act and Assert section comments.
---

# Test conventions

For every new or modified test in this project:

- Always add `// Act` immediately before the operation being tested.
- Always add `// Assert` immediately before checking the result. Saving and reopening a presentation to inspect the result belongs in this section.
- Use the exact comment spelling `// Act` and `// Assert`, matching the existing C# test style.
- Keep setup before the Act section; mark it with `// Arrange` for consistency with the project's tests.
- Before finishing, check that every test added or modified has both required comments and that each comment marks the correct section.

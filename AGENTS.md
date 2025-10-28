# AGENTS.md - AI Agent Guide for ShapeCrawler

This document provides guidance for AI coding assistants working on the ShapeCrawler project.

## Project Overview

ShapeCrawler is a .NET library that provides a simplified object model for manipulating PowerPoint presentations. It wraps the Open XML SDK to offer a more intuitive API for processing `.pptx` files without requiring Microsoft Office installation.

**Core Purpose**: Simplify PowerPoint presentation manipulation through an object-oriented API that abstracts away the complexity of the Open XML format.

## Architecture

### Project Structure

```
src/                        # Main source code
tests/
├── ShapeCrawler.DevTests/  # Fast developer tests to check for regression changes
└── ShapeCrawler.CITests/   # Slow tests that are generally only run on GitHub Workflow along with fast developer tests

```

### Key Design Patterns

1. **Interface-based API**: Public API is exposed through interfaces (e.g., `IPresentation`, `ISlide`, `IShape`)
2. **Wrapper Pattern**: Internal classes wrap Open XML SDK elements
3. **Object-Oriented Design**: Emphasis on nouns for class names, no static members in classes
4. **Encapsulation**: Logic is encapsulated in constructors and internal/public methods

## Code Style Guidelines

### Mandatory Rules
- **What is an Object?** The project follows the principle that the correct object is a representation of a real-world entity or concept. In its constructor, the class encapsulates properties or another object as “coordinates” that the class instance will use to refer to the real-world entity.
- **Naming Conventions**:
    - Class names must be **nouns** (e.g., `Slide`, `Slides`, not `SlideManager` or `SlideService`)
    - No `-er`, `-or`, `-service` suffixes
- **Method complexity**:
  - The maximum allowed method Cognitive Complexity is 15.
  - The maximum allowed method Cyclomatic Complexity is 10.
- **File Size Limit**: Keep files under 500 lines. If a file exceeds this, extract logic into new classes/files.

- **Instance Members**: Use `this` prefix for all instance members
   ```csharp
   // Good
   this.fieldName = value;
   
   // Bad
   fieldName = value;

- **Code comments:** 
  - Use WHY comment instead of WHAT comment
    ```csharp
    // Good
    var cTitle = cChart?.Title;
    if (cTitle == null)
    { 
          return 18; // used by the PowerPoint application as the default font size 
    }
    
    // Bad
    var cTitle = cChart?.Title;
    if (cTitle == null)
    { 
          return 18; // default font size 
    }
   
  - Use "Open XML", not "OpenXML".

- **No Public/Internal Static Members**: Classes should not have public or internal static members. Encapsulate behavior in instance methods.
- **File-Scoped Namespaces**: Always use file-scoped namespace declarations
   ```csharp
   namespace ShapeCrawler.Charts;
   
   public class Chart { }
   ```
- **Primary Constructors**: Prefer primary constructors (C# 12+) where applicable
- **XML Documentation**: All public and internal members must have XML documentation comments

### EditorConfig Enforcement

The project uses strict `.editorconfig` rules. Key settings:
- **Indentation**: 4 spaces
- **Line Endings**: CRLF
- **Nullable Reference Types**: Enabled and strictly enforced
- **StyleCop**: Extensive StyleCop analyzers enabled
- **Build**: Release configuration enforces all code style rules

## Testing Guidelines

### Test Project
- **Always use**: `tests/ShapeCrawler.DevTests/ShapeCrawler.DevTests.csproj`
- **Never use**: Other test projects (they're for CI/CD)

### Test Requirements
- **Side-Effect Tests**: Tests that modify presentations must call `.Validate()` in assertions
- **Quantity**: Write only ONE test when asked unless explicitly requested otherwise

### Running Tests
```bash
dotnet test tests/ShapeCrawler.DevTests/ShapeCrawler.DevTests.csproj
```

## Common Workflows

### Adding New Features
1. Identify the appropriate namespace/folder (e.g., `Shapes/`, `Charts/`)
2. Create interface first (if public API)
3. Implement internal class
4. Keep files under 500 lines
5. Add XML documentation
6. Write test
7. Build in Release configuration

### Bug Fixes
1. Locate the issue in the codebase
2. Write a failing test that reproduces the bug
3. Fix the bug
4. Ensure test passes and `.Validate()` is called if side effects exist
5. Build in Release configuration

### Making Changes
1. Read relevant files to understand context
2. Make targeted changes
3. Run tests: `dotnet test tests/ShapeCrawler.DevTests/ShapeCrawler.DevTests.csproj`
4. Build Release: `dotnet build src/ShapeCrawler.csproj -c Release`
5. Fix any linter errors

## Build and Validation

### Development Build
```bash
# Debug configuration - lenient for development
dotnet build src/ShapeCrawler.csproj -c Debug
```

### Release Build
```bash
# Release configuration - enforces all style rules
dotnet build src/ShapeCrawler.csproj -c Release
```

**Critical**: Always build in Release configuration before completing work to ensure all code style checkers pass.

## Common Pitfalls to Avoid

- ❌ **Don't create Manager/Service/Helper classes**
   - Use noun-based classes with instance methods

- ❌ **Don't exceed 500 lines per file**
   - Extract into new files/classes

- ❌ **Don't skip `this` prefix**
   - Always use for instance members

- ❌ **Don't forget XML documentation**
   - Required for all public/internal members

- ❌ **Don't use static members in classes**
   - Encapsulate in instance methods

- ❌ **Don't test with wrong project**
   - Only use `ShapeCrawler.DevTests`

- ❌ **Don't skip Release build**
   - Required to catch all linter/style issues

## Useful Context

### Typical User Operations
- Load/create presentations
- Access slides and shapes
- Manipulate text, images, tables, charts
- Save modifications

### API Design Philosophy
- Fluent and intuitive
- Hide Open XML complexity
- Null-safe with nullable reference types
- Interface-based for testability

### Example Code Patterns
```csharp
// Loading and accessing
var pres = new Presentation("file.pptx");
var shape = pres.Slide(1).Shapes.Shape("TextBox 1");
var text = shape.TextBox!.Text;

// Creating
var pres = new Presentation(p => p.Slide());
pres.Slide(1).Shapes.AddShape(x: 50, y: 60, width: 100, height: 70);
pres.Save("output.pptx");
```

## Quick Reference

| Task | Command |
|------|---------|
| Run tests | `dotnet test tests/ShapeCrawler.DevTests/ShapeCrawler.DevTests.csproj` |
| Build (dev) | `dotnet build src/ShapeCrawler.csproj -c Debug` |
| Build (release) | `dotnet build src/ShapeCrawler.csproj -c Release` |

## Resources

- **Project Issues**: [GitHub Issues](https://github.com/ShapeCrawler/ShapeCrawler/issues)
- **Discussions**: [GitHub Discussions](https://github.com/ShapeCrawler/ShapeCrawler/discussions)
- **Examples**: See `examples/` folder for usage patterns


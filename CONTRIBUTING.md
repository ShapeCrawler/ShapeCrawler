# Contributing to ShapeCrawler
A big welcome and thank you for considering contributing to ShapeCrawler's open source project!

Reading and following these guidelines will help us make the contribution process easy and effective for everyone involved. It also communicates that you agree to respect the time of the developers managing and developing these open source project. In return, we will reciprocate that respect by addressing your issues, assessing changes, and helping you finalize your pull requests.

## Development Flow
1. Fork the repository
2. Create a branch
3. Open solution `ShapeCrawler.Dev.sln` in your favorite IDE
4. Make a feature or fix bug
    - use Debug configuration for development to avoid early code style issues during building
5. Test
    - code changes always should be covered with tests
    - the tests that test changes with side effect must have in the assertion block calling `.Validate()` for the presentations 
6. **Build the Release configuration of the solution to ensure that all code style checkers pass**
7. Create a Pull Request

## Recommended tools
The internal structure of PowerPoint presentation is one of the most difficult among other Microsoft Office documents. For example, the slide presented for a user is not a single document but only top on Slide Layout and Slide Master slides. Even just that levels frequently lead to confusion. The following is a list of tools and notes that should simplify your development while working on ShapeCrawler's issue.

- **[OOXML Viewer](https://marketplace.visualstudio.com/items?itemName=yuenm18.ooxml-viewer)** — extension for Visual Studio Code allowing to view Open XML package. One of the good features of this extension is track changes of modified presentation.
- **PPTX files is ZIP archives** — rename `.pptx` to `.zip` to inspect.
# Developer guidelines
The [OpenXML specification](https://www.ecma-international.org/publications/standards/Ecma-376.htm) is a large and complicated beast. In order for PptxXML, the wrapper around OpenXML, to support all the features, I rely on community contributions.

Here are some rules and tips.

* All branches are based on `develop` and have followed pattern name `feature/ISS-{Issue number}`, regardless of a feature or bug type. For example, *feature/ISS-3*.
* Using *Rebase* over *Merge* to get remote `develop` changes is preferable.
* All draft (temporary) branches are named `draft/{branch name}`.
* Pull request is submitted to `develop` branch.
* Where possible, pull requests should include unit tests that cover as many uses cases as possible.
* Attache problem presentation file to creating issue as possible.
---
## Working with PowerPoint file internals

PowerPoint file (`.pptx`) is zip package. You can easily verify this by renaming the extension any PowerPoint file to `.zip` and opening the file in your favourite `.zip` file editor.

---
## Code conventions
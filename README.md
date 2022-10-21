
<h3 align="center">

![ShapeCrawler](./resources/readme.png)

</h3>

<h3 align="center">

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=orange)](https://www.nuget.org/packages/ShapeCrawler) ![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange) [![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE) 

</h3>

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) to process presentations without having Microsoft Office installed.

## Getting Started

> `install-package ShapeCrawler`

### Usage

```c#
using var pres = SCPresentation.Open("some.pptx");

// get number of slides
var slidesCount = pres.Slides.Count;

// get text of TextBox 
var autoShape = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1");
var text = autoShape.TextFrame!.Text;
```

Visit [**Wiki**](https://github.com/ShapeCrawler/ShapeCrawler/wiki/Examples) page to find more usage samples.

# Have questions?

If you have a question:
- [join](https://github.com/ShapeCrawler/ShapeCrawler/discussions/categories/q-a) our Discussions Forum  and open discussion;
- you can always email the author to theadamo86@gmail.com

# Contributing
How you can contribute?
- **Give a Star**⭐ If you find this useful, please give it a star to show your support.
- **Polls**. Participate in the voting on [Polls](https://github.com/ShapeCrawler/ShapeCrawler/discussions/categories/polls) discussion.
- **Bug report**. If you get some issue, please don't ignore and report this bug on [issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) page.
- **Code contributing**. There are features/bugs tagged with [help-wanted](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aissue+is%3Aopen+label%3A%22help+wanted%22) label which waiting for your Pull Request🙂 Please read [Contribution Guide](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CONTRIBUTING.md) to get more details.

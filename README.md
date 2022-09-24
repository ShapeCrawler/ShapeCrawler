
<h3 align="center">

![ShapeCrawler](./resources/readme.png)

</h3>

<h3 align="center">

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=orange)](https://www.nuget.org/packages/ShapeCrawler) ![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange) [![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE) 

</h3>

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) to process presentations without having Microsoft Office installed.

## Getting Started

### Install

To get started, install ShapeCrawler from [NuGet](https://nuget.org/packages/ShapeCrawler):

```console
dotnet add package ShapeCrawler
```

The library currently supports the following frameworks: 
- .NET 5+ 
- .NET Core 2.0+
- .NET Framework 4.6.1+

### Usage

```c#
using var pres = SCPresentation.Open("helloWorld.pptx", isEditable: false);
var slidesCount = pres.Slides.Count;
var autoShape = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1");
Console.WriteLine(autoShape.TextBox.Text);
```

Visit [**Wiki**](https://github.com/ShapeCrawler/ShapeCrawler/wiki#examples) page to find more usage samples.

# Have questions?

If you have a question:
- [join](https://github.com/ShapeCrawler/ShapeCrawler/discussions/categories/q-a) our Discussions Forum  and open discussion;
- you can always email the author to theadamo86@gmail.com

# Contributing
How you can contribute?
- **Give a Star**⭐ If you find this useful, please give it a star to show your support.
- **Polls**. Participate in the voting on [Polls](https://github.com/ShapeCrawler/ShapeCrawler/discussions/categories/polls) discussion.
- **Bug report**. If you get some issue, please don't ignore and report the bug on [issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) page.
- **Implement feature**. [Some features/bugs](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aissue+is%3Aopen+label%3A%22help+wanted%22) are tagged with *help-wanted* label and waiting for your Pull Request🙂 Please visit [Contribution Guide](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aissue+is%3Aopen+label%3A%22help+wanted%22) to get some development recommendations.
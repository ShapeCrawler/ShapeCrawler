<h3 align="center">

![ShapeCrawler](./docs/logo-extend.png)

</h3>

<h3 align="center"> 

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=orange)](https://www.nuget.org/packages/ShapeCrawler) [![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?color=orange)](https://makeapullrequest.com) ![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange) [![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE) 

</h3>

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), allowing users to process presentations without having Microsoft Office installed.

## Contents

- [Quick Start](#quick-start)
- [How To?](#how-to)
  - [Create presentation](#create-presentation)
  - [More samples](#more-samples)
- [Prerelease Version](#prerelease-version)
- [Have questions?](#have-questions)
- [How to contribute](#how-to-contribute)
  - [Bug Report](#bug-report)
  - [Code Contributing](#code-contributing)

## Quick Start
> `install-package ShapeCrawler`

```c#
// open existing presentation
var pres = new Presentation("some.pptx");

var shapes = pres.Slides[0].Shapes;

// get number of shapes on slide
var shapesCount = shapes.Count;

// get text
var shape = shapes.GetByName("TextBox 1");
var text = shape.TextFrame!.Text;
```

## How To?

### Create presentation

```c#
// create a new presentation
var pres = new Presentation();

var shapes = pres.Slides[0].Shapes;

// add new shape
shapes.AddRectangle(x: 50, y: 60, width: 100, height: 70);
var addedShape = shapes.Last();

addedShape.TextFrame!.Text = "Hello World!";

pres.SaveAs("my_pres.pptx");
```

### More samples

Visit the [**Wiki**](https://github.com/ShapeCrawler/ShapeCrawler/wiki/Examples) page to find more usage samples.

## Prerelease Version
To access prerelease builds from `master` branch, add `https://www.myget.org/F/shape/api/v3/index.json` as a package source:

![Prerelease](./docs/prerelease.png)
![Download Prerelease](./docs/prerelease-download.png)

## Have questions?

If you have a question:
- [Join](https://github.com/ShapeCrawler/ShapeCrawler/discussions/categories/q-a) our Discussions Forum  and open a discussion;
- You can always email the author at theadamo86@gmail.com

## How to contribute?
Give a star⭐ if you find this useful, please give it a star to show your support.

### Bug Report
If you encounter an issue, report the bug on the [issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) page.

To be able to reproduce a bug, it's often necessary to have the original presentation file attached to the issue description. If this file contains confidential data and cannot be shared publicly, you can securely send it to theadamo86@gmail.com. Of course, if your security policy allow this. We assure you that only the maintainer will access this file, and it will not be shared publicly.

### Code contributing
Pull Requests are welcome! Please read the [Contribution Guide](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CONTRIBUTING.md) for more details.

## Changelog  

### Version 0.50.4 - 2024-05-10
🐞Fixed `ISlideShapes.AddPicture()` [#671](https://github.com/ShapeCrawler/ShapeCrawler/issues/671)

Visit [CHANGELOG.md](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md) to see the full log.
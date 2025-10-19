<h3 align="center">

![ShapeCrawler](./assets/logo.png)

</h3>

<h3 align="center"> 

[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?color=orange)](https://makeapullrequest.com)
![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange)
[![GitHub repo Good Issues for newbies](https://img.shields.io/github/issues/ShapeCrawler/ShapeCrawler/good%20first%20issue?style=flat&logo=github&logoColor=green&label=Good%20First%20issues)](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aopen+is%3Aissue+label%3A%22good+first+issue%22)
[![GitHub Help Wanted issues](https://img.shields.io/github/issues/ShapeCrawler/ShapeCrawler/help%20wanted?style=flat&logo=github&logoColor=b545d1&label=%22Help%20Wanted%22%20issues)](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aopen+is%3Aissue+label%3A%22help+wanted%22)

</h3>

ShapeCrawler is a .NET library for manipulating PowerPoint presentations. It provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), allowing users to process presentations without having Microsoft Office installed.

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
> `dotnet add package ShapeCrawler`

```C#
var pres = new Presentation("pres.pptx");
var shapes = pres.Slide(1).Shapes;

var shapesCount = shapes.Count;

// Get text
var shape = shapes.Shape("TextBox 1");
var text = shape.TextBox!.Text;
```

## How To?

### Create presentation

```C#
var pres = new Presentation(p => p.Slide());

var shapes = pres.Slide(1).Shapes;

shapes.AddShape(x: 50, y: 60, width: 100, height: 70);
var addedShape = shapes.Last();

addedShape.TextBox!.SetText("Hello World!");

pres.Save("pres.pptx");
```

### Update picture
```C#
var pres = new Presentation("pres.pptx");
var picture = pres.Slide(1).Shape("Picture 1").Picture!;

var image = System.IO.File.OpenRead("new-image.png");
picture.Image!.Update(image);
pres.Save();

var mimeType = picture.Image!.Mime;
```

### More samples
You can find more usage samples in [**Examples**](https://github.com/ShapeCrawler/ShapeCrawler/tree/master/examples).

## Prerelease Version
To access the latest prerelease builds from the branch `master`, use the NuGet package source `https://www.myget.org/F/shape/api/v3/index.json`.

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

### Version 0.75.0 - 2025-10-19
🍀Added support for updating the font size of chart title [#1135](https://github.com/ShapeCrawler/ShapeCrawler/issues/1135)

Visit [CHANGELOG.md](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md) to see the full change history.
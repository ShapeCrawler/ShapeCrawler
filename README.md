
<h3 align="center">

![ShapeCrawler](./resources/readme.png)

</h3>

<h3 align="center">

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=orange)](https://www.nuget.org/packages/ShapeCrawler) ![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange) [![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE) 

</h3>

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), allowing users to process presentations without having Microsoft Office installed.

## Getting Started

> `install-package ShapeCrawler`

### Usage

#### Read presentation

```c#
// open existing presentation
using var pres = SCPresentation.Open("some.pptx");

var shapeCollection = pres.Slides[0].Shapes;

// get number of shapes on slide
var slidesCount = shapeCollection.Count;

// get text
var autoShape = shapeCollection.GetByName<IAutoShape>("TextBox 1");
var text = autoShape.TextFrame!.Text;
```

#### Create presentation

```c#
// create a new presentation
var pres = SCPresentation.Create();

var shapeCollection = pres.Slides[0].Shapes;

// add new shape
var addedShape = shapeCollection.AddAutoShape(SCAutoShapeType.TextBox, x: 50, y: 60, width: 100, height: 70);

addedShape.TextFrame!.Text = "Hello World!";

pres.SaveAs("my_pres.pptx");
```

### More samples

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

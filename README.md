<h3 align="center">

![ShapeCrawler](./doc/logo-extend.png)

</h3>

<h3 align="center"> 

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=orange)](https://www.nuget.org/packages/ShapeCrawler) ![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange) [![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE) 

</h3>

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), allowing users to process presentations without having Microsoft Office installed.


⚠️**Warning:** Since 15 February, the library collects usage data to help us to improve your experience. It is collected by the maintainer and not shared with the community. You can opt out of telemetry. For more details, please visit [Statistics Collection](https://github.com/ShapeCrawler/ShapeCrawler#statistics-collection).

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
var addedShape = shapeCollection.AddRectangle(x: 50, y: 60, w: 100, h: 70);

addedShape.TextFrame!.Text = "Hello World!";

pres.SaveAs("my_pres.pptx");
```

### More samples

Visit [**Wiki**](https://github.com/ShapeCrawler/ShapeCrawler/wiki/Examples) page to find more usage samples.

## Have questions?

If you have a question:
- [join](https://github.com/ShapeCrawler/ShapeCrawler/discussions/categories/q-a) our Discussions Forum  and open discussion;
- you can always email the author to theadamo86@gmail.com

## Contributing
How you can contribute?
- **Give a Star**⭐ If you find this useful, please give it a star to show your support.
- **Bug report**. If you get some issue, please don't ignore and report this bug on [issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) page.
- **Code contributing**. There are features/bugs tagged with [help-wanted](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aissue+is%3Aopen+label%3A%22help+wanted%22) label which waiting for your Pull Request🙂 Please read [Contribution Guide](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CONTRIBUTING.md) to get more details.


## Changelog
## Version 0.43.0 - 2023-04-06
🍀Added `IShapeCollection.AddLine()` to add Line shape [#465](https://github.com/ShapeCrawler/ShapeCrawler/issues/465)

Visit [CHANGELOG.md](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md) to see the full log.

## Statistics Collection

Since 15 February, the library collects usage data to help us to improve your experience. It is collected by the maintainer and not shared with the community. Rest assured that we do not collect any sensitive or presentation content data. The collection will include, for example, information on the operating system, target framework, and frequently used shape types being utilized. If you prefer not to participate in this data collection, you can easily opt-out by setting the global setting `SCSettings.CanCollectLogs = false`.
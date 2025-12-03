<h3 align="center">

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="./assets/logo-dark.png">
  <source media="(prefers-color-scheme: light)" srcset="./assets/logo.png">
  <img alt="ShapeCrawler" src="./assets/logo.png">
</picture>

</h3>

<h3 align="center"> 

[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?color=orange)](https://makeapullrequest.com)
![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange)
[![GitHub repo Good Issues for newbies](https://img.shields.io/github/issues/ShapeCrawler/ShapeCrawler/good%20first%20issue?style=flat&logo=github&logoColor=green&label=Good%20First%20issues)](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aopen+is%3Aissue+label%3A%22good+first+issue%22)
[![GitHub Help Wanted issues](https://img.shields.io/github/issues/ShapeCrawler/ShapeCrawler/help%20wanted?style=flat&logo=github&logoColor=b545d1&label=%22Help%20Wanted%22%20issues)](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aopen+is%3Aissue+label%3A%22help+wanted%22)

</h3>

<p align="center">
  <strong>A .NET library for manipulating PowerPoint presentations without Microsoft Office</strong>
</p>

ShapeCrawler provides a clean, intuitive API on top of the Open XML SDK, making it easy to read, create, and modify `.pptx` files programmatically.

---

## 📦 Installation

```bash
dotnet add package ShapeCrawler
```

## 🚀 Getting Started

```csharp
// Load an existing presentation
var pres = new Presentation("presentation.pptx");

// Access shapes on a slide
var shapes = pres.Slide(1).Shapes;
var textBox = shapes.Shape("TextBox 1");

// Read text content
var text = textBox.TextBox!.Text;

// Modify and save
textBox.TextBox!.SetText("Updated content");
pres.Save();
```

## 🎯 Why ShapeCrawler?

- **No Office Required** – Process presentations on any platform without Microsoft Office installation
- **Clean API** – Intuitive object model that hides the complexity of Open XML

## 💡 Common Use Cases

### Creating Presentations from Scratch

```csharp
// Create a new presentation with a slide
var pres = new Presentation(p => p.Slide());

// Add a shape with text
var shapes = pres.Slide(1).Shapes;
shapes.AddShape(x: 50, y: 60, width: 100, height: 70);

var addedShape = shapes.Last();
addedShape.TextBox!.SetText("Hello World!");

pres.Save("output.pptx");
```

### Updating Images

```csharp
var pres = new Presentation("presentation.pptx");
var picture = pres.Slide(1).Shape("Picture 1").Picture!;

// Replace the image
using var newImage = File.OpenRead("new-image.png");
picture.Image!.Update(newImage);

pres.Save();
```

### Line

#### Adding a straight line

```csharp
var pres = new Presentation("presentation.pptx");
var shapes = pres.Slide(1).Shapes;

// Add a line from (50, 60) to (100, 60)
shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);
```

#### Accessing Start and End Points

```csharp
var pres = new Presentation("presentation.pptx");
var line = pres.Slide(1).Shapes.First(shape => shape.GeometryType == Geometry.Line).Line;

var start = line.StartPoint; // Point(x, y)
var end = line.EndPoint;     // Point(x, y)
Console.WriteLine($"Line from {start.X},{start.Y} to {end.X},{end.Y}");
```

### More Examples

**[See More Examples](https://github.com/ShapeCrawler/ShapeCrawler/tree/master/examples)**

## 🌟 Getting Help

Have questions? We're here to help!

- [**Issues**](https://github.com/ShapeCrawler/ShapeCrawler/issues) – Report bugs or request features
- [**Discussions Forum**](https://github.com/ShapeCrawler/ShapeCrawler/discussions) – Ask questions and share ideas
- Email – Reach out to theadamo86@gmail.com

## 🤝 Contributing

We love contributions! Here's how you can help:

- Give us a star ⭐ – If you find ShapeCrawler useful, show your support with a star!
- Reporting Bugs – Found a bug? [Open an issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) with a clear description of the problem
- Need to share a confidential file? – Email it to theadamo86@gmail.com – only the maintainer will access it
- Contribute Code – Pull requests are welcome! Check out our:
  - [**Good First Issues**](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aopen+is%3Aissue+label%3A%22good+first+issue%22) – Perfect for newcomers
  - [**Contribution Guide**](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CONTRIBUTING.md) – Guidelines and best practices

## 🔄 Pre-release Versions

Want to try the latest features? Access pre-release builds from the `master` branch:

**NuGet Feed:** `https://www.myget.org/F/shape/api/v3/index.json`

## 📝 Changelog

### Version 0.76.1 - 2025-12-03
🐞Fixed potential security vulnerabilities

[**View Full Changelog**](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md)

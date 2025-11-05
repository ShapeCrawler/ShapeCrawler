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

## 🚀 Why ShapeCrawler?

- **No Office Required** – Process presentations on any platform without Microsoft Office installation
- **Clean API** – Intuitive object model that hides the complexity of Open XML
- **Production Ready** – Battle-tested in real-world applications with comprehensive test coverage
- **Actively Maintained** – Regular updates and responsive to community feedback

## 📦 Installation

```bash
dotnet add package ShapeCrawler
```

## 🎯 Quick Start

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

### Working with Tables and Charts

ShapeCrawler supports comprehensive manipulation of:
- **Tables** – Create, modify cells, styling
- **Charts** – Update data, titles, formatting
- **Text** – Rich text formatting, fonts, paragraphs
- **Media** – Images, audio, video

**[📚 See More Examples](https://github.com/ShapeCrawler/ShapeCrawler/tree/master/examples)**

## 🔧 Advanced Features

- Full shape manipulation (position, size, rotation, styling)
- Table operations (add/remove rows/columns, cell merging)
- Chart data and formatting updates
- Text and paragraph formatting
- Slide master and layout access
- Image cropping and replacement
- Embedded media handling

## 🌟 Getting Help

**Have questions?** We're here to help!

- 💬 [**Discussions Forum**](https://github.com/ShapeCrawler/ShapeCrawler/discussions) – Ask questions and share ideas
- 📧 **Email** – Reach out to theadamo86@gmail.com
- 🐛 [**Issues**](https://github.com/ShapeCrawler/ShapeCrawler/issues) – Report bugs or request features

## 🤝 Contributing

We love contributions! Here's how you can help:

**⭐ Give us a star** – If you find ShapeCrawler useful, show your support with a star!

### Reporting Bugs

Found a bug? [Open an issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) with:
- A clear description of the problem
- Steps to reproduce
- Expected vs. actual behavior

**Need to share a confidential file?** Email it to theadamo86@gmail.com – only the maintainer will access it.

### Contributing Code

Pull requests are welcome! Check out our:
- [**Good First Issues**](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aopen+is%3Aissue+label%3A%22good+first+issue%22) – Perfect for newcomers
- [**Contribution Guide**](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CONTRIBUTING.md) – Guidelines and best practices

## 🔄 Pre-release Versions

Want to try the latest features? Access pre-release builds from the `master` branch:

**NuGet Feed:** `https://www.myget.org/F/shape/api/v3/index.json`

## 📝 Changelog

### Version 0.75.2 - 2025-11-05
🐞Fixed adding a slide with an image background slide layout [#1156](https://github.com/ShapeCrawler/ShapeCrawler/issues/1156)

[**View Full Changelog**](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md)

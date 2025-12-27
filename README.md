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
[![GitHub Help Wanted issues](https://img.shields.io/github/issues/ShapeCrawler/ShapeCrawler/help%20wanted?style=flat&logo=github&logoColor=b545d1&label=%22Help%20Wanted%22%20issues)](https://github.com/ShapeCrawler/ShapeCrawler/issues?q=is%3Aopen+is%3Aissue+label%3A%22help+wanted%22)

</h3>

<p align="center">
  <strong>PowerPoint (PPTX) manipulation library for .NET / C# developers</strong>
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
var text = textBox.TextBox.Text;

// Modify and save
textBox.TextBox.SetText("Updated content");
pres.Save();
```

## 🎯 Why ShapeCrawler?

- **No Office Required** – Process presentations on any platform without Microsoft Office installation
- **Clean API** – Intuitive object model that hides the complexity of Open XML
- **Open Source** — Actively maintained

## 💡 Common Use Cases

### Create presentations

```csharp
// Create a new presentation with a slide
var pres = new Presentation(p => p.Slide());

// Add a shape with text
var shapes = pres.Slide(1).Shapes;
shapes.AddShape(x: 50, y: 60, width: 100, height: 70);

var addedShape = shapes.Last();
addedShape.TextBox.SetText("Hello World!");

pres.Save("output.pptx");
```

### Update image

```csharp
var pres = new Presentation("presentation.pptx");
var picture = pres.Slide(1).Shape("Picture 1").Picture;

// Replace the image
using var newImage = File.OpenRead("new-image.png");
picture.Image.Update(newImage);

pres.Save();
```

### Tables

#### Create table

```csharp
var pres = new Presentation("presentation.pptx");
var shapes = pres.Slide(1).Shapes;

// Add a 3x2 table at position (50, 120)
shapes.AddTable(x: 50, y: 120, columnsCount: 3, rowsCount: 2);

var table = shapes.Last().Table;
table[0, 0].TextBox.SetText("Hello table");

pres.Save();
```

#### Update table

```csharp
var pres = new Presentation("presentation.pptx");
var table = pres.Slide(1).Shapes.Shape("Table 1").Table;

// Insert a row at index 1, using row 0 as a template
table.Rows.Add(1, 0);

// Merge two header cells
table.MergeCells(table[0, 0], table[0, 1]);

pres.Save();
```

### Lines

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

### Charts

#### Create Bar Chart

```csharp
var pres = new Presentation(p => p.Slide());
var shapes = pres.Slide(1).Shapes;

var points = new Dictionary<string, double>
{
    { "Q1", 50 },
    { "Q2", 60 },
    { "Q3", 40 }
};

// Add a bar chart
shapes.AddBarChart(x: 100, y: 100, width: 500, height: 350, points, "Sales");

pres.Save("output.pptx");
```

#### Update Chart Category

```csharp
var pres = new Presentation("presentation.pptx");
var chart = pres.Slide(1).Shapes.Shape("Bar Chart 1").Chart;

// Update category name
chart.Categories[0].Name = "Renamed Category";

pres.Save();
```

### More Examples

**[See More Examples](https://github.com/ShapeCrawler/ShapeCrawler/tree/master/examples)**

## ❓ Getting Help

Have questions? We're here to help!

- [Issues](https://github.com/ShapeCrawler/ShapeCrawler/issues) – Report bugs or request features
- [Discussions Forum](https://github.com/ShapeCrawler/ShapeCrawler/discussions) – Ask questions and share ideas
- Email – Reach out to theadamo86@gmail.com

## 🤝 Contributing

We love contributions! Here's how you can help:

- Give us a star ⭐ – If you find ShapeCrawler useful, show your support with a star!
- Reporting Bugs – Found a bug? [Open an issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) with a clear description of the problem
- Contribute Code – Pull requests are welcome!
- Need to share a confidential file? – Email it to theadamo86@gmail.com – only the maintainer will access it

## 🔄 Pre-release Versions

Want to try the latest features? Access pre-release builds from the `master` branch using the following NuGet: `https://www.myget.org/F/shape/api/v3/index.json`

## 📝 Changelog

### Version 0.77.0 - 2025-12-24
🍀Added support for updating the category name of the multi-category chart [#151](https://github.com/ShapeCrawler/ShapeCrawler/issues/151)

[**View Full Changelog**](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md)

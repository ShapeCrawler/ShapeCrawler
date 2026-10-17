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

Requires .NET 10 or later.

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

### Read a fill's theme color name

```csharp
var pres = new Presentation("presentation.pptx");
var fill = pres.Slide(1).Shape("Rectangle 1").Fill;

// For example, "Text 1" or "Accent 2"; null for fills without an explicit theme color.
var themeColorName = fill.ThemeColorName;
```

The same property is available on a table cell's `Fill`. Brightness and transparency adjustments preserve the theme color name.

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
var chart = pres.Slide(1).Shapes.Shape("Bar Chart 1").BarChart;

// Update category name
chart.Categories[0].Name = "Renamed Category";

pres.Save();
```

#### Add Chart Category

```csharp
var pres = new Presentation("presentation.pptx");
var chart = pres.Slide(1).Shapes.Shape("Bar Chart 1").BarChart;

// Supply one value for each series, in series collection order.
chart.Categories!.Add("Q4", 70, 80, 90);

pres.Save();
```

Single-level text categories are supported in non-combination charts. For charts linked
to an embedded Excel workbook, category and value ranges must be vertical; the workbook
and chart caches are updated together.

#### Add Chart Data Point

```csharp
var pres = new Presentation("presentation.pptx");
var chart = pres.Slide(1).Shapes.Shape("Scatter Chart 1").ScatterChart;
var series = chart.SeriesCollection[0];

series.XPoints!.Add(6);
series.Points.Add(12.5);

pres.Save();
```

`Points`, `XPoints`, and `BubbleSizePoints` support appending numeric values to cached
or literal data. Append matching coordinates and bubble sizes separately. Formula-backed
data requires an embedded Excel workbook, a contiguous vertical range, and an unoccupied
next cell. The chart cache, formula, and workbook are updated together. Use `Categories.Add`
when adding a category and values for every series.

### More Examples

**[See More Examples](https://github.com/ShapeCrawler/ShapeCrawler/tree/master/examples)**

## ❓ Getting Help

Have questions? We're here to help!

- [Issues](https://github.com/ShapeCrawler/ShapeCrawler/issues) – Report bugs or request features
- Email – Reach out to theadamo86@gmail.com

## 🤝 Contributing

We love contributions! Here's how you can help:

- Give us a star ⭐ – If you find ShapeCrawler useful, show your support with a star!
- Reporting Bugs – Found a bug? [Open an issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) with a clear description of the problem
- Need to share a confidential file? – Email it to theadamo86@gmail.com – only the maintainer will access it

## 🔄 Pre-release Versions

Want to try the latest features? Access pre-release builds from the `master` branch using the following NuGet: `https://www.myget.org/F/shape/api/v3/index.json`

## 📝 Changelog

### Version 0.81.0 - 2026-10-17
🍀Added `ISeriesCollection.Add()` to add series to bar/column, line, pie/doughnut, scatter, and area charts [#490](https://github.com/ShapeCrawler/ShapeCrawler/issues/490)  
🍀Added support for creating and reading all PowerPoint line types and updating line endpoints [#217](https://github.com/ShapeCrawler/ShapeCrawler/issues/217)  
🍀Added APIs to update SVG pictures and MP3/WAVE audio content [#362](https://github.com/ShapeCrawler/ShapeCrawler/issues/362), [#248](https://github.com/ShapeCrawler/ShapeCrawler/issues/248)

[**View Full Changelog**](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md)

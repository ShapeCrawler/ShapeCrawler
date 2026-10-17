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

### Add Markdown as editable text

```csharp
using var pres = new Presentation(p => p.Slide());
pres.Slide(1).Shapes.AddMarkdownTextBox(50, 60, 500, 300, """
    # Project update

    **Completed** work and *next steps*: [details](https://example.com).

    - First item
      - Nested item
    - Second item
    """);
pres.Save("markdown.pptx");
```

Coordinates and dimensions are in points. The result is a regular PowerPoint text box:
use `Shapes.Last().TextBox` to edit its text, paragraphs, and portions after adding it.
Content stays in the supplied box; it is not paginated across slides.

Markdown is parsed with [Markdig](https://github.com/xoofx/markdig) using its default CommonMark syntax.

| Markdown | PowerPoint result |
| --- | --- |
| Bold, italic, nested emphasis | Formatted text portions |
| Inline links, reference links, `<https://…>` and email autolinks | Clickable text with hyperlink relationships |
| Paragraphs and headings | Separate paragraphs; headings are bold at the default text size |
| Bulleted and numbered lists | Native bullets and numbering, including the starting number; nested indentation is capped at nine levels |
| Multiple paragraphs inside a list item | Continuation paragraphs aligned with the item text, without another bullet |
| Soft line breaks / hard line breaks | A space / a line break within the same paragraph |
| Inline and fenced/indented code | Literal text in Consolas; code block line breaks are preserved |
| Images | Alt text only; no image is downloaded or embedded |
| Block quotes | Their paragraphs and inline formatting, without quote-specific styling |
| Thematic breaks (`---`) | Omitted |
| Raw HTML | Literal tags and content; no HTML rendering, CSS, or script execution |
| Extension syntax (tables, strikethrough, task checkboxes, math, etc.) | Extensions are disabled; syntax remains literal text within otherwise supported blocks |

Blank lines separate paragraphs without creating extra empty paragraphs. Empty Markdown creates an empty text box;
`null` throws `ArgumentNullException`. Numbered lists use decimal numbers followed by periods; starting numbers
outside PowerPoint's range are clamped to 1–32767. Link titles and code language hints are ignored.
This API imports Markdown; it does not import HTML documents or reproduce browser layout.

### Embed Word, Excel, or PowerPoint documents

```csharp
using var pres = new Presentation(p => p.Slide());
using var document = File.OpenRead("report.docx");
using var preview = File.OpenRead("report-preview.png");
pres.Slide(1).Shapes.AddOleObject(50, 60, 300, 200, document, OleObjectType.WordDocument, preview);
pres.Save("with-document.pptx");
```

Use `ExcelWorkbook` for `.xlsx` or `PowerPointPresentation` for `.pptx`. Coordinates and dimensions are in points.
Supply a preview image to display on the slide; ShapeCrawler embeds the original document without rendering it.
Both streams are read from their current positions and remain open. Opening or editing the embedded document requires a compatible Office application.

To insert an Excel object on the first slide of an existing presentation (PowerPoint's **Insert > Object > Create from file**):

```csharp
using var pres = new Presentation("presentation.pptx");
using var workbook = File.OpenRead("workbook.xlsx");
using var preview = File.OpenRead("workbook-preview.png");
pres.Slide(1).Shapes.AddOleObject(50, 60, 300, 200, workbook, OleObjectType.ExcelWorkbook, preview);
pres.Save("with-workbook.pptx");
```

### Set audio timing

```csharp
using var pres = new Presentation(p => p.Slide());
var shapes = pres.Slide(1).Shapes;
using var audio = File.OpenRead("audio.mp3");
shapes.AddAudio(100, 100, audio);
var media = shapes.Last().Media!;
media.StartMode = AudioStartMode.Automatically; // With Previous in animation Timing
media.StartDelay = TimeSpan.FromSeconds(2);
pres.Save("audio-timing.pptx");
```

`StartDelay` defaults to zero and is saved with millisecond precision. Changing `StartMode` preserves the delay.

### Organize slides into sections

```csharp
using var pres = new Presentation("presentation.pptx");
pres.Sections.Add("Introduction", firstSlideNumber: 1);
var content = pres.Sections.Add("Content", firstSlideNumber: 3);
pres.Save("with-sections.pptx");
```

Slide numbers start at 1. Each new section contains slides from its starting slide up to the next section, or the end of the presentation. Adding a section splits the existing section at that slide. When the first section starts after slide 1, earlier slides are placed in a `Default Section`.

### Remove a shape's fill

```csharp
using var pres = new Presentation("presentation.pptx");
pres.Slide(1).Shape("Rectangle 1").Fill.SetNoFill();
pres.Save("no-fill.pptx");
```

`SetNoFill()` replaces the existing fill and preserves the outline. To remove fills from shapes inside a group, call `SetNoFill()` on each member's `Fill` through `GroupedShapes`; repeat for members of nested groups.

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

### Read and update picture alternative text

```csharp
using var pres = new Presentation("presentation.pptx");
var pictureShape = pres.Slide(1).Shape("Picture 1");

var altText = pictureShape.AltText;
pictureShape.AltText = "A mountain lake at sunrise";
pres.Save("with-alt-text.pptx");
```

`AltText` belongs to the shape and is stored in the picture's `descr` attribute. It returns an empty string when no alternative text is set. Assign an empty string to clear it.

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

#### Rename Chart Series (Column Title)

```csharp
var pres = new Presentation("presentation.pptx");
var chart = pres.Slide(1).Shapes.Shape("Bar Chart 1").BarChart!;

chart.SeriesCollection[0].Name = "Cats";
chart.SeriesCollection[1].Name = "Dogs";

pres.Save();
```

Assign `ISeries.Name` to update the series title shown in the chart legend.
For charts linked to an embedded Excel workbook, the corresponding header cell
is updated together with the chart cache. Use `SeriesCollection` in the current API
in place of the older `SeriesList` property.

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

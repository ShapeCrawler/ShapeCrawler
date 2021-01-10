<h3 align="center">

![ShapeCrawler](/resources/readme.png)

</h3>

<h3 align="center">

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=blue)](https://www.nuget.org/packages/ShapeCrawler) [![.NET Standard](https://img.shields.io/badge/.NET%20Core-2.0-blue)](#) [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-blue.svg)](#) [![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE) 

</h3>

✅ **Project status: active**

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides fluent APIs to process slides without having Microsoft Office installed.

## Getting Started
You can quickly start work with the library by following steps listed below.
## Install

- [NuGet](https://nuget.org/packages/ShapeCrawler): `dotnet add package ShapeCrawler`

## Usage
### Work with Text
```C#
PresentationEx presentation = PresentationEx.Open("helloWorld.pptx", true);
SlideEx slide = presentation.Slides.First();

// Prints on console content of all text shapes
foreach (ShapeEx shape in slide.Shapes)
{
    if (shape.HasTextFrame)
    {
        Console.WriteLine(shape.TextFrame.Text);
    }
}

// Changes paragraph text
ShapeEx textShape = slide.Shapes.First(sp => sp.HasTextFrame);
ParagraphEx paragraph = textShape.TextFrame.Paragraphs.First();
paragraph.Text = "A new paragraph text";
presentation.Save();
```
### Work with Chart
```C#
PresentationEx presentation = PresentationEx.Open("helloWorld.pptx", isEditable: false);
SlideEx slide = presentation.Slides.First();
ShapeEx chartShape = slide.Shapes.First(sp => sp.HasChart == true);

IChart chart = chartShape.Chart;
if (chart.HasTitle)
{
    Debug.Print(chart.Title);
}
if (chart.Type == ChartType.BarChart)
{
    Debug.Print("Chart type is BarChart.");
}
```
# Feedback and Give a Star! :star:
The project is in development, and I’m pretty sure there are still lots of things to add in this library. Try it out and let me know your thoughts.

Feel free to submit a [ticket](https://github.com/ShapeCrawler/ShapeCrawler/issues) if you find bugs. Your valuable feedback is much appreciated to better improve this project. If you find this useful, please give it a star to show your support for this project. 

# Contributing
1. Fork it (https://github.com/ShapeCrawler/ShapeCrawler/fork)
2. Create your feature branch (`git checkout -b my-new-feature`) from master.
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin my-new-feature`).
5. Create a new Pull Request.

# Changelog
## Version 0.10.0 - 2021-01-01
### Added
- Added `Portion.Remove()` to be able to remove paragraph portion;
- Added setter for `Paragraph.Text` property to be able to change paragraph's text;
- Added support for .NET Core 2.0

To find out more, please check out the [CHANGELOG](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md).

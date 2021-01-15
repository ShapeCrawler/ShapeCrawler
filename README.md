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

### Open presentation
```C#
PresentationSc presentation = PresentationSc.Open(@"c:\MyPresentations\helloWorld.pptx", isEditable: false);

Console.WriteLine($"Number of slides in the presentation: {presentation.Slides.Count}");
Console.WriteLine($"Number of shapes on the first slide: {presentation.Slides[0].Shapes.Count}");

presentation.Close();
```

### Work with Text shape
```C#
PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", true);
SlideSc slide = presentation.Slides.First();

// Prints on console content of all text shapes
foreach (ShapeSc shape in slide.Shapes)
{
    if (shape.HasTextFrame)
    {
        Console.WriteLine(shape.TextFrame.Text);
    }
}

// Changes paragraph text
ShapeSc textShape = slide.Shapes.First(sp => sp.HasTextFrame);
Paragraph paragraph = textShape.TextFrame.Paragraphs.First();
paragraph.Text = "A new paragraph text";
presentation.Save();
```

#### Font
```C#
PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", false);
SlideSc slide = presentation.Slides.First();

// Prints font name and sizes of paragraph text portions
ITextFrame textFrame = slide.Shapes.First(sp => sp.HasTextFrame).TextFrame;
IEnumerable<Portion> paragraphPortions = textFrame.Paragraphs.First().Portions;
foreach (Portion portion in paragraphPortions)
{
    Console.WriteLine($"Font name: {portion.Font.Name}");
    Console.WriteLine($"Font size: {portion.Font.Size}");
}
```
### Work with Chart shape
```C#
PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: false);
SlideSc slide = presentation.Slides.First();
ShapeSc chartShape = slide.Shapes.First(sp => sp.HasChart == true);

ChartSc chart = chartShape.Chart;
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
## Version 0.11.0 - 2021-01-10
### Added
- Added setter for `Portion.Text` property to be able to change text of paragraph portion (#22)
- Added setter for `Portion.Font.Name` to change font name of the portion of non-placeholder shape (#82)
- Added setter for `Portion.Font.Size` to change font size of the portion of non-placeholer shape (#81)

To find out more, please check out the [CHANGELOG](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md).

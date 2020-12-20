<h3 align="center">

![ShapeCrawler](/resources/readme.png)

</h3>

<h3 align="center">

  [![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
  [![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=blue)](https://www.nuget.org/packages/ShapeCrawler)  

</h3>

ShapeCrawler (formerly SlideDotNet) is a fluent API for the processing of PowerPoint presentations without Microsoft Office installed.

## Getting Started
You can quickly start work with the library by following steps listed below.

### Installing
To install ShapeCrawler, run the following command in the Package Manager Console:
```
PM> Install-Package ShapeCrawler
```

### Usage

```C#
public static async void Usage()
{
    var presentation = new PresentationEx(@"c:\test.pptx");
    var slides = presentation.Slides;
    var numSlides = slides.Count();

    // Gets slide sizes in EMUs
    int slideHeight = presentation.SlideHeight;
    int slideWidth = presentation.SlideWidth;

    // Saves presentation
    presentation.SaveAs(@"c:\test_edited.pptx");

    // Gets number of shapes
    Slide slide = slides[0];
    var shapes = slide.Shapes;
    var numShapes = shapes.Count;

    // Gets slide number
    int slideNumber = slide.Number;

    // Gets slide background content
    byte[] backgroundBytes = await slide.BackgroundImage.GetImageBytes();
}
```
<details>
<summary><i>Show more usage examples...</i></summary>

```C#
public static async void Usage()
{
    // Gets number of slides
    var presentation = new PresentationEx(@"c:\test.pptx");
    var slides = presentation.Slides;
    var numSlides = slides.Count();

    // Gets slide sizes in EMUs
    int slideHeight = presentation.SlideHeight;
    int slideWidth = presentation.SlideWidth;

    // Saves presentation
    presentation.SaveAs(@"c:\test_edited.pptx");

    // Gets number of shapes
    Slide slide = slides[0];
    var shapes = slide.Shapes;
    var numShapes = shapes.Count;

    // Gets slide number
    int slideNumber = slide.Number;

    // Gets slide background content
    byte[] backgroundBytes = await slide.BackgroundImage.GetImageBytes();

    // Sets slide background
    using (FileStream fs = File.OpenRead(@"c:\test.png"))
    {
        slide.BackgroundImage.SetImageStream(fs);
    }

    // Set some custom data in slide, e.g. tag
    slide.CustomData = "#mySlide";

    // Prints texts of shapes on the Debug console
    foreach (var sp in shapes)
    {
        if (sp.HasTextFrame)
        {
            Debug.WriteLine(sp.TextFrame.Text);
        }
    }

    // Works with charts
    var chartShape = shapes.FirstOrDefault(s => s.HasChart);
    if (chartShape != null)
    {
        IChart chart = chartShape.Chart;
        if (chart.HasTitle)
        {
            Debug.Print(chart.Title);
        }
        if (chart.Type == ChartType.BarChart)
        {
            Debug.Print("Chart type is BarChart.");
        }
    }
}
```
</details>

## Changelog
### Version 0.7.0 - 2020-10-12
#### Added
- Added `Bullet` property for the paragraph:
    - Bullet.Type
    - Bullet.Char
    - Bullet.FontName
    - Bullet.Size
    - Bullet.ColorHex

To find out more, please check out the [CHANGELOG](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md).

## Contribution
- Feel free to report a bug or suggest a new feature by creating an [issue](https://github.com/ShapeCrawler/ShapeCrawler/issues);
- Welcome to contribute. We are wating for your [Pull Requests](https://github.com/ShapeCrawler/ShapeCrawler/pulls). 
